import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(page_title="Incidencias / Ausentismo / Asistencia", layout="wide")
st.title("App Incidencias / Ausentismo / Asistencia")

# =========================
# Constantes / opciones
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
}
CLASIF_OPTS = ["Seleccionar", "Injustificada", "Permiso", "No Procede - Cambio de Turno"]

# =========================
# Helpers
# =========================
def normalize_rut(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = s.replace(".", "").replace(" ", "")
    return s

def try_parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def excel_to_df(file, sheet_index=0):
    return pd.read_excel(file, sheet_name=sheet_index, engine="openpyxl")

def find_col(df: pd.DataFrame, candidates):
    norm_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = str(cand).strip().lower()
        if key in norm_map:
            return norm_map[key]
    return None

def get_num(df: pd.DataFrame, colname_candidates):
    col = find_col(df, colname_candidates if isinstance(colname_candidates, list) else [colname_candidates])
    if not col:
        return pd.Series([0.0] * len(df))
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)

def safe_text_series(df: pd.DataFrame, candidates, default=""):
    col = find_col(df, candidates)
    if not col:
        return pd.Series([default] * len(df))
    return df[col].astype(str).fillna(default)

def split_full_name(full: str):
    """Heurística: últimos 2 tokens = apellidos; resto = nombres."""
    if full is None or (isinstance(full, float) and pd.isna(full)):
        return "", "", ""
    s = str(full).strip()
    if not s:
        return "", "", ""
    parts = [p for p in s.split() if p]
    if len(parts) >= 3:
        return " ".join(parts[:-2]), parts[-2], parts[-1]
    if len(parts) == 2:
        return parts[0], parts[1], ""
    return parts[0], "", ""

# =========================
# Excel export (Cabify + dropdown)
# =========================
THIN = Side(style="thin", color="D0D0D0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def style_ws_cabify(ws):
    header_fill = PatternFill("solid", fgColor=CABIFY["m2"])
    header_font = Font(color="FFFFFF", bold=True)
    center = Alignment(vertical="center", horizontal="center", wrap_text=True)
    left = Alignment(vertical="top", horizontal="left", wrap_text=True)

    ws.freeze_panes = "A2"

    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = BORDER

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.border = BORDER
            cell.alignment = left

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col[:2000]:
            try:
                v = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(v))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(10, max_len + 2), 45)

def write_df_to_sheet(wb: Workbook, sheet_name: str, df: pd.DataFrame):
    ws = wb.create_sheet(title=sheet_name[:31])
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    style_ws_cabify(ws)
    return ws

def apply_dropdown(ws, headers, col_name, options):
    if col_name not in headers:
        return
    col_idx = list(headers).index(col_name) + 1
    start_row = 2
    end_row = max(2, ws.max_row)
    col_letter = ws.cell(row=1, column=col_idx).column_letter

    dv = DataValidation(
        type="list",
        formula1='"{}"'.format(",".join(options)),
        allow_blank=False,
        showDropDown=True
    )
    dv.error = "Selecciona una opción válida."
    dv.prompt = "Selecciona una clasificación."
    dv.add(f"{col_letter}{start_row}:{col_letter}{end_row}")
    ws.add_data_validation(dv)

def set_date_format(ws, headers, col_name, fmt="dd/mm/yyyy"):
    if col_name not in headers:
        return
    col_idx = list(headers).index(col_name) + 1
    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=col_idx).number_format = fmt

def set_header_row_date_format(ws, start_col_idx=2, fmt="dd/mm/yyyy"):
    """Para hojas tipo matriz donde la fila 1 (desde columna B) son fechas."""
    for c in range(start_col_idx, ws.max_column + 1):
        ws.cell(row=1, column=c).number_format = fmt

def to_excel_bytes(dfs: dict):
    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    for name, df in dfs.items():
        ws = write_df_to_sheet(wb, name, df)

        if name == "Incidencias":
            apply_dropdown(ws, df.columns, "Clasificación Manual", CLASIF_OPTS)
            set_date_format(ws, df.columns, "Fecha")

        if name == "KPIs_Diarios":
            set_header_row_date_format(ws, start_col_idx=2, fmt="dd/mm/yyyy")

        if "Fecha" in df.columns and name != "Incidencias":
            set_date_format(ws, df.columns, "Fecha")

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
df_turnos = excel_to_df(f_turnos, 0)  # por ahora no se usa, queda para futuras reglas
df_activos = excel_to_df(f_reporte_turnos, 0)

df_inasist = excel_to_df(f_detalle, 0)  # Hoja 1
df_asist = excel_to_df(f_detalle, 1)    # Hoja 2

# =========================
# Normalización: RUT + Fecha base
# =========================
rut_col_inas = find_col(df_inasist, ["RUT", "Rut", "rut"])
rut_col_as = find_col(df_asist, ["RUT", "Rut", "rut"])
if not rut_col_inas or not rut_col_as:
    st.error("No pude detectar la columna RUT en una de las hojas del 'Detalle Turnos Colaboradores'.")
    st.stop()

df_inasist["RUT_norm"] = df_inasist[rut_col_inas].apply(normalize_rut)
df_asist["RUT_norm"] = df_asist[rut_col_as].apply(normalize_rut)

dia_col_inas = find_col(df_inasist, ["Día", "Dia", "DIA", "día"])
df_inasist["Fecha_base"] = df_inasist[dia_col_inas].apply(try_parse_date_any) if dia_col_inas else pd.NaT

fecha_ent_col = find_col(df_asist, ["Fecha Entrada", "Fecha entrada", "FECHA ENTRADA"])
dia_col_as = find_col(df_asist, ["Día", "Dia", "DIA", "día"])
if fecha_ent_col:
    df_asist["Fecha_base"] = df_asist[fecha_ent_col].apply(try_parse_date_any)
elif dia_col_as:
    df_asist["Fecha_base"] = df_asist[dia_col_as].apply(try_parse_date_any)
else:
    df_asist["Fecha_base"] = pd.NaT

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
# (NUEVO) Filtro de colaboradores válidos:
# Solo considerar RUTs presentes en Detalle Turnos Colaboradores
# =========================
allowed_ruts = set(
    pd.concat([df_inasist["RUT_norm"], df_asist["RUT_norm"]], ignore_index=True)
      .dropna()
      .astype(str)
      .str.strip()
)
allowed_ruts.discard("")
if not allowed_ruts:
    st.warning("No pude construir la lista de colaboradores válidos desde el Detalle Turnos Colaboradores (RUTs vacíos).")

# =========================
# Turnos planificados (Activos + Turnos) -> largo
# =========================
if "RUT" not in df_activos.columns:
    rut_col_act = find_col(df_activos, ["RUT", "Rut", "rut"])
    if rut_col_act and rut_col_act != "RUT":
        df_activos = df_activos.rename(columns={rut_col_act: "RUT"})

df_activos["RUT_norm"] = df_activos["RUT"].apply(normalize_rut) if "RUT" in df_activos.columns else ""

if allowed_ruts:
    df_activos = df_activos[df_activos["RUT_norm"].isin(allowed_ruts)].copy()

fixed_cols = [c for c in ["Nombre del Colaborador", "RUT", "Área", "Supervisor"] if c in df_activos.columns]
date_cols = [c for c in df_activos.columns if c not in fixed_cols + ["RUT_norm"]]

df_act_long = df_activos.melt(
    id_vars=[c for c in fixed_cols if c in df_activos.columns] + (["RUT_norm"] if "RUT_norm" in df_activos.columns else []),
    value_vars=date_cols,
    var_name="Fecha",
    value_name="Turno_planificado"
)
df_act_long["Fecha_dt"] = df_act_long["Fecha"].apply(try_parse_date_any)
df_act_long["Turno_planificado"] = df_act_long["Turno_planificado"].astype(str).str.strip()
df_act_long.loc[df_act_long["Turno_planificado"].isin(["", "nan", "NaT", "None", "-", "L", "l"]), "Turno_planificado"] = ""

# =========================
# Construcción Incidencias (campos solicitados)
# - Asistencias: SOLO si hay Retraso o Salida Anticipada
# - Inasistencias: todas (se clasifica manualmente)
# =========================
inc_rows = []

retr = get_num(df_asist, ["Retraso (horas)", "Retraso horas", "Retraso"])
sal = get_num(df_asist, ["Salida Anticipada (horas)", "Salida Anticipada", "Salida anticipada (horas)"])
mask_asist = (retr > 0) | (sal > 0)

df_asist_inc = df_asist[mask_asist].copy()
df_asist_inc["Fecha"] = pd.to_datetime(df_asist_inc["Fecha_base"], errors="coerce").dt.date
df_asist_inc["Nombre"] = safe_text_series(df_asist_inc, ["Nombre"], "")
df_asist_inc["Primer Apellido"] = safe_text_series(df_asist_inc, ["Primer Apellido", "Primer apellido"], "")
df_asist_inc["Segundo Apellido"] = safe_text_series(df_asist_inc, ["Segundo Apellido", "Segundo apellido"], "")
df_asist_inc["RUT"] = df_asist_inc["RUT_norm"].astype(str)

df_asist_inc["Turno"] = safe_text_series(df_asist_inc, ["Turno"], "")
df_asist_inc["Especialidad"] = safe_text_series(df_asist_inc, ["Especialidad"], "")
df_asist_inc["Supervisor"] = safe_text_series(df_asist_inc, ["Supervisor"], "")

df_asist_inc["Tipo_Incidencia"] = "Incidencia"
df_asist_inc["Detalle"] = "Retraso_h=" + retr[mask_asist].astype(str) + " | SalidaAnt_h=" + sal[mask_asist].astype(str)
df_asist_inc["Clasificación Manual"] = "Seleccionar"

inc_rows.append(df_asist_inc[[
    "Fecha", "Nombre", "Primer Apellido", "Segundo Apellido", "RUT",
    "Turno", "Especialidad", "Supervisor",
    "Tipo_Incidencia", "Detalle", "Clasificación Manual"
]])

df_inasist_inc = df_inasist.copy()
df_inasist_inc["Fecha"] = pd.to_datetime(df_inasist_inc["Fecha_base"], errors="coerce").dt.date
df_inasist_inc["Nombre"] = safe_text_series(df_inasist_inc, ["Nombre"], "")
df_inasist_inc["Primer Apellido"] = safe_text_series(df_inasist_inc, ["Primer Apellido", "Primer apellido"], "")
df_inasist_inc["Segundo Apellido"] = safe_text_series(df_inasist_inc, ["Segundo Apellido", "Segundo apellido"], "")
df_inasist_inc["RUT"] = df_inasist_inc["RUT_norm"].astype(str)

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
# Cumplimiento por colaborador (dinámico)
# =========================
st.subheader("Cumplimiento por colaborador (dinámico)")

df_act_long_valid = df_act_long[df_act_long["Turno_planificado"] != ""].copy()

turnos_plan = (
    df_act_long_valid.groupby("RUT_norm")
    .size()
    .reset_index(name="Turnos_planificados")
)

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

if "Nombre del Colaborador" in df_activos.columns:
    tmp = df_activos[["RUT_norm", "Nombre del Colaborador"]].dropna(subset=["RUT_norm"]).drop_duplicates("RUT_norm").copy()
    n1, ap1, ap2 = [], [], []
    for v in tmp["Nombre del Colaborador"].tolist():
        a, b, c = split_full_name(v)
        n1.append(a); ap1.append(b); ap2.append(c)
    tmp["Nombre"] = n1
    tmp["Primer Apellido"] = ap1
    tmp["Segundo Apellido"] = ap2
    name_map = tmp[["RUT_norm", "Nombre", "Primer Apellido", "Segundo Apellido"]]
else:
    name_map = (
        edited_tmp.dropna(subset=["RUT_norm"])
        .drop_duplicates("RUT_norm")[["RUT_norm", "Nombre", "Primer Apellido", "Segundo Apellido"]]
    )

cumpl = cumpl.merge(name_map, on="RUT_norm", how="left")

cumpl = cumpl[[
    "Nombre", "Primer Apellido", "Segundo Apellido", "RUT_norm",
    "Turnos_planificados", "Injustificadas", "Cumplimiento_%"
]].rename(columns={"RUT_norm": "RUT"}).sort_values("Cumplimiento_%", ascending=True)

st.dataframe(cumpl, use_container_width=True)

# =========================
# (NUEVO) KPIs diarios (matriz)
# =========================
st.subheader("KPIs diarios (matriz)")

planned_day = (
    df_act_long_valid.dropna(subset=["Fecha_dt"])
    .groupby(df_act_long_valid["Fecha_dt"].dt.date)
    .size()
    .rename("Turnos_planificados")
    .reset_index(names="Fecha")
)

df_inj2 = edited_tmp[edited_tmp["Clasificación Manual"] == "Injustificada"].copy()
df_inj2["Fecha_d"] = pd.to_datetime(df_inj2["Fecha"], errors="coerce").dt.date

inj_total_day = df_inj2.groupby("Fecha_d").size().rename("Injustificadas_total").reset_index()

inj_by_tipo = (
    df_inj2.groupby(["Fecha_d", "Tipo_Incidencia"])
    .size()
    .unstack(fill_value=0)
    .reset_index()
)
if "Inasistencia" not in inj_by_tipo.columns:
    inj_by_tipo["Inasistencia"] = 0
if "Incidencia" not in inj_by_tipo.columns:
    inj_by_tipo["Incidencia"] = 0
inj_by_tipo = inj_by_tipo.rename(columns={"Fecha_d": "Fecha", "Inasistencia": "Inasistencias_inj", "Incidencia": "Incidencias_inj"})

kpi_day = planned_day.merge(inj_total_day, left_on="Fecha", right_on="Fecha_d", how="left").drop(columns=["Fecha_d"])
kpi_day = kpi_day.merge(inj_by_tipo, on="Fecha", how="left")
kpi_day["Injustificadas_total"] = kpi_day["Injustificadas_total"].fillna(0).astype(int)
kpi_day["Inasistencias_inj"] = kpi_day["Inasistencias_inj"].fillna(0).astype(int)
kpi_day["Incidencias_inj"] = kpi_day["Incidencias_inj"].fillna(0).astype(int)

kpi_day["Cumplimiento_%"] = (1 - (kpi_day["Injustificadas_total"] / kpi_day["Turnos_planificados"])).clip(lower=0, upper=1) * 100

kpi_labels = [
    ("Turnos planificados", "Turnos_planificados"),
    ("Inasistencias injustificadas", "Inasistencias_inj"),
    ("Incidencias injustificadas", "Incidencias_inj"),
    ("Total injustificadas", "Injustificadas_total"),
    ("% cumplimiento", "Cumplimiento_%"),
]

kpi_day = kpi_day.sort_values("Fecha")
dates = kpi_day["Fecha"].tolist()

matrix = {"KPI": [lab for lab, _ in kpi_labels]}
for d in dates:
    row = kpi_day[kpi_day["Fecha"] == d].iloc[0]
    matrix[d] = [row[key] for _, key in kpi_labels]

kpis_diarios = pd.DataFrame(matrix)
st.dataframe(kpis_diarios, use_container_width=True)

# =========================
# Export
# =========================
st.subheader("Descarga")

edited_export = edited.copy()
edited_export["Fecha"] = pd.to_datetime(edited_export["Fecha"], errors="coerce")

excel_bytes = to_excel_bytes({
    "Incidencias": edited_export,
    "Resumen_Injustificadas": resumen,
    "Cumplimiento": cumpl,
    "KPIs_Diarios": kpis_diarios
})

st.download_button(
    "Descargar Excel consolidado (Cabify + dropdown)",
    data=excel_bytes,
    file_name="reporte_incidencias_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

