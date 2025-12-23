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
# Config
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
    "green": "0C936B",  # tu input decía #OC936B (O letra). Debe ser 0 (cero)
}

CLASIF_OPTS = ["Seleccionar", "Injustificada", "Permiso", "No Procede - Cambio de Turno"]

# =========================
# Helpers
# =========================
def normalize_rut(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().upper().replace(".", "").replace(" ", "")

def try_parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def excel_to_df(file, sheet_index=0):
    return pd.read_excel(file, sheet_name=sheet_index, engine="openpyxl")

def find_col(df: pd.DataFrame, candidates):
    norm_map = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        k = str(cand).strip().lower()
        if k in norm_map:
            return norm_map[k]
    # match suave
    for cand in candidates:
        k = str(cand).strip().lower()
        for kk, real in norm_map.items():
            if kk == k:
                return real
    return None

def get_num(df, candidates):
    col = find_col(df, candidates if isinstance(candidates, list) else [candidates])
    if not col:
        return pd.Series([0.0] * len(df))
    return pd.to_numeric(df[col], errors="coerce").fillna(0.0)

def safe_text_series(df, candidates, default=""):
    col = find_col(df, candidates)
    if not col:
        return pd.Series([default] * len(df))
    return df[col].astype(str).fillna(default)

def maybe_filter_area(df, only_area_value):
    if not only_area_value:
        return df
    area_col = find_col(df, ["Área", "Area", "AREA"])
    if not area_col:
        return df
    return df[df[area_col].astype(str).str.upper().str.contains(str(only_area_value).upper(), na=False)].copy()

def split_fullname(fullname: str):
    """
    Intenta separar: Nombre(s) + 1er Apellido + 2do Apellido
    Regla simple: últimos 2 tokens = apellidos, resto = nombre(s)
    """
    if not fullname or pd.isna(fullname):
        return "", "", ""
    toks = str(fullname).strip().split()
    if len(toks) == 1:
        return toks[0], "", ""
    if len(toks) == 2:
        return toks[0], toks[1], ""
    nombre = " ".join(toks[:-2])
    ap1 = toks[-2]
    ap2 = toks[-1]
    return nombre, ap1, ap2

# =========================
# Excel styling + dropdown
# =========================
def style_ws_cabify(ws):
    header_fill = PatternFill("solid", fgColor=CABIFY["m4"])
    header_font = Font(color="FFFFFF", bold=True)
    border_side = Side(style="thin", color=CABIFY["m2"])
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    alt_fill = PatternFill("solid", fgColor=CABIFY["m10"])
    base_fill = PatternFill("solid", fgColor=CABIFY["m11"])

    # header
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # body (alternado)
    for r in range(2, ws.max_row + 1):
        fill = alt_fill if (r % 2 == 0) else base_fill
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    # autofilter
    if ws.max_column >= 1 and ws.max_row >= 1:
        ws.auto_filter.ref = ws.dimensions

    # ancho columnas (simple)
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = 18

def set_date_format(ws, df_cols, col_name="Fecha"):
    if col_name not in df_cols:
        return
    idx = list(df_cols).index(col_name) + 1
    for r in range(2, ws.max_row + 1):
        c = ws.cell(row=r, column=idx)
        c.number_format = "dd-mm-yyyy"

def write_df_to_sheet(wb, name, df: pd.DataFrame):
    ws = wb.create_sheet(title=name[:31])
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    return ws

def ensure_list_sheet(wb):
    if "Listas" in wb.sheetnames:
        return wb["Listas"]
    ws = wb.create_sheet("Listas")
    ws.append(["Clasificación Manual"])
    for opt in CLASIF_OPTS:
        ws.append([opt])
    ws.column_dimensions["A"].width = 30
    return ws

def apply_dropdown(ws, df_cols, target_col_name="Clasificación Manual"):
    if target_col_name not in df_cols:
        return
    col_idx = list(df_cols).index(target_col_name) + 1
    col_letter = ws.cell(row=1, column=col_idx).column_letter

    # Validación referenciando Listas!$A$2:$A$5 (sin header)
    dv = DataValidation(type="list", formula1="=Listas!$A$2:$A$5", allow_blank=False)
    dv.error = "Selecciona una opción válida"
    dv.errorTitle = "Opción inválida"
    dv.prompt = "Selecciona una opción"
    dv.promptTitle = "Clasificación Manual"

    dv.add(f"{col_letter}2:{col_letter}1048576")
    ws.add_data_validation(dv)

def to_excel_bytes(dfs: dict, dropdown_sheet_name="Incidencias"):
    output = BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    # sheet listas para dropdown
    ws_list = ensure_list_sheet(wb)
    ws_list.sheet_state = "hidden"  # oculto, pero existe para validación

    for name, df in dfs.items():
        ws = write_df_to_sheet(wb, name, df)

        # Fecha formateada
        if "Fecha" in df.columns:
            set_date_format(ws, df.columns, "Fecha")

        # Dropdown solo en hoja principal
        if name == dropdown_sheet_name and "Clasificación Manual" in df.columns:
            apply_dropdown(ws, df.columns, "Clasificación Manual")

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
    st.subheader("Filtros")
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    min_inc_h = st.number_input("Tiempo mínimo incidencia (horas)", min_value=0.0, value=0.0, step=0.25)
    st.caption("Se considera incidencia si Retraso ≥ umbral o Salida Anticipada ≥ umbral o (Retraso+Salida) ≥ umbral.")

if not all([f_turnos, f_reporte_turnos, f_detalle]):
    st.info("Sube los 3 archivos para comenzar.")
    st.stop()

# =========================
# Load
# =========================
df_turnos = excel_to_df(f_turnos, 0)  # por ahora no se usa, queda listo para reglas futuras
df_activos = excel_to_df(f_reporte_turnos, 0)

df_inasist = excel_to_df(f_detalle, 0)  # Hoja 1
df_asist = excel_to_df(f_detalle, 1)    # Hoja 2

# Detectar RUT en detalle
rut_col_inas = find_col(df_inasist, ["RUT", "Rut", "rut"])
rut_col_as = find_col(df_asist, ["RUT", "Rut", "rut"])
if not rut_col_inas or not rut_col_as:
    st.error("No pude detectar la columna RUT en una de las hojas del Detalle Turnos Colaboradores.")
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

# Filtrar por área (opcional)
df_inasist = maybe_filter_area(df_inasist, only_area)
df_asist = maybe_filter_area(df_asist, only_area)
df_activos = maybe_filter_area(df_activos, only_area)

# =========================
# Turnos planificados (reporte turnos) -> formato largo
# =========================
# columnas fijas típicas
fixed_cols_candidates = ["Nombre del Colaborador", "RUT", "Área", "Supervisor"]
fixed_cols = [c for c in fixed_cols_candidates if c in df_activos.columns]
date_cols = [c for c in df_activos.columns if c not in fixed_cols]

# asegurar columna RUT
if "RUT" not in df_activos.columns:
    rut_col_act = find_col(df_activos, ["RUT", "Rut", "rut"])
    if rut_col_act:
        df_activos = df_activos.rename(columns={rut_col_act: "RUT"})
        if "RUT" not in fixed_cols:
            fixed_cols = [c for c in fixed_cols_candidates if c in df_activos.columns]
        date_cols = [c for c in df_activos.columns if c not in fixed_cols]

df_act_long = df_activos.melt(
    id_vars=[c for c in fixed_cols if c in df_activos.columns],
    value_vars=date_cols,
    var_name="Fecha",
    value_name="Turno_planificado"
)

df_act_long["Fecha_dt"] = df_act_long["Fecha"].apply(try_parse_date_any)
df_act_long["RUT_norm"] = df_act_long["RUT"].apply(normalize_rut) if "RUT" in df_act_long.columns else ""

df_act_long["Turno_planificado"] = df_act_long["Turno_planificado"].astype(str).str.strip()
df_act_long.loc[df_act_long["Turno_planificado"].isin(["", "nan", "NaT", "None", "-", "—"]), "Turno_planificado"] = ""

# excluir libres (L) para planificación (tu regla)
df_act_long["Turno_planificado_clean"] = df_act_long["Turno_planificado"].copy()
df_act_long.loc[df_act_long["Turno_planificado_clean"].str.upper().isin(["L", "LIBRE"]), "Turno_planificado_clean"] = ""

# =========================
# Selector de fechas (se mantiene)
# =========================
min_date_candidates = []
for s in [df_act_long["Fecha_dt"], df_inasist["Fecha_base"], df_asist["Fecha_base"]]:
    s_ok = s.dropna()
    if len(s_ok):
        min_date_candidates.append(s_ok.min())
        min_date_candidates.append(s_ok.max())

if not min_date_candidates:
    st.error("No pude detectar fechas válidas en los archivos.")
    st.stop()

fecha_min = min(min_date_candidates)
fecha_max = max(min_date_candidates)

fecha_desde, fecha_hasta = st.date_input(
    "📅 Selecciona rango de fechas:",
    value=(fecha_min.date(), fecha_max.date())
)

# filtrar por rango
def filter_by_range(df, col):
    if col not in df.columns:
        return df
    s = pd.to_datetime(df[col], errors="coerce")
    return df[(s.dt.date >= fecha_desde) & (s.dt.date <= fecha_hasta)].copy()

df_inasist = filter_by_range(df_inasist, "Fecha_base")
df_asist = filter_by_range(df_asist, "Fecha_base")
df_act_long = df_act_long[(df_act_long["Fecha_dt"].dt.date >= fecha_desde) & (df_act_long["Fecha_dt"].dt.date <= fecha_hasta)].copy()

# =========================
# (1) Filtrar colaboradores: SOLO los que existan en Detalle Turnos Colaboradores
# =========================
valid_ruts = set(pd.concat([df_inasist["RUT_norm"], df_asist["RUT_norm"]], ignore_index=True).dropna().unique().tolist())
df_act_long = df_act_long[df_act_long["RUT_norm"].isin(valid_ruts)].copy()
df_inasist = df_inasist[df_inasist["RUT_norm"].isin(valid_ruts)].copy()
df_asist = df_asist[df_asist["RUT_norm"].isin(valid_ruts)].copy()

# =========================
# Incidencias: Asistencias (solo si supera umbral) + Inasistencias (para clasificar)
# =========================
inc_rows = []

# Asistencias: retraso / salida anticipada (con umbral)
retr = get_num(df_asist, ["Retraso (horas)", "Retraso horas", "Retraso"])
sal = get_num(df_asist, ["Salida Anticipada (horas)", "Salida Anticipada", "Salida anticipada (horas)"])
total_rs = retr + sal
umbral = float(min_inc_h)

mask_asist = (retr >= umbral) | (sal >= umbral) | (total_rs >= umbral)
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
    + " | Total_h=" + total_rs[mask_asist].astype(str).values
)
df_asist_inc["Clasificación Manual"] = "Seleccionar"

inc_rows.append(df_asist_inc[[
    "Fecha", "Nombre", "Primer Apellido", "Segundo Apellido", "RUT",
    "Turno", "Especialidad", "Supervisor",
    "Tipo_Incidencia", "Detalle", "Clasificación Manual"
]])

# Inasistencias: se listan completas (del rango) para clasificar
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

# Orden y fecha
df_incidencias["Fecha"] = pd.to_datetime(df_incidencias["Fecha"], errors="coerce")
df_incidencias = df_incidencias.sort_values(["Fecha", "RUT"], na_position="last").reset_index(drop=True)

# =========================
# UI principal
# =========================
st.subheader("Reporte Total de Incidencias (para clasificar)")

edited = st.data_editor(
    df_incidencias,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Clasificación Manual": st.column_config.SelectboxColumn(
            options=CLASIF_OPTS
        )
    }
)

# =========================
# Resumen dinámico (se actualiza cuando editas)
# =========================
st.subheader("Resumen dinámico (según Clasificación Manual)")

resumen = (
    edited.groupby(["Clasificación Manual", "Tipo_Incidencia"], dropna=False)
    .size()
    .reset_index()
    .rename(columns={0: "Cantidad"})
    .sort_values("Cantidad", ascending=False)
)
st.dataframe(resumen, use_container_width=True)

# =========================
# Cumplimiento por colaborador (base = turnos planificados activos sin 'L')
# =========================
st.subheader("Cumplimiento por colaborador (base = turnos planificados del periodo, sin Libres)")

df_turnos_valid = df_act_long[df_act_long["Turno_planificado_clean"] != ""].copy()

# turnos planificados por rut (en el rango filtrado)
turnos_plan = (
    df_turnos_valid.groupby("RUT_norm")
    .size()
    .reset_index(name="Turnos_planificados")
)

# injustificadas por rut (desde la tabla editada)
tmp = edited.copy()
tmp["RUT_norm"] = tmp["RUT"].apply(normalize_rut)
inj_cnt = (
    tmp[tmp["Clasificación Manual"] == "Injustificada"]
    .groupby("RUT_norm")
    .size()
    .reset_index(name="Injustificadas")
)

cumpl = turnos_plan.merge(inj_cnt, on="RUT_norm", how="left")
cumpl["Injustificadas"] = cumpl["Injustificadas"].fillna(0).astype(int)

# nombres: preferir reporte turnos (Nombre del Colaborador)
name_col = find_col(df_activos, ["Nombre del Colaborador", "Nombre", "Colaborador"])
if name_col:
    base_names = df_activos.copy()
    base_names["RUT_norm"] = base_names["RUT"].apply(normalize_rut)
    base_names = base_names[base_names["RUT_norm"].isin(valid_ruts)].copy()

    base_names = base_names.drop_duplicates("RUT_norm")[["RUT_norm", name_col]]
    base_names[["Nombre", "Primer Apellido", "Segundo Apellido"]] = base_names[name_col].apply(
        lambda x: pd.Series(split_fullname(x))
    )
else:
    base_names = tmp.drop_duplicates("RUT_norm")[["RUT_norm", "Nombre", "Primer Apellido", "Segundo Apellido"]]

cumpl = cumpl.merge(base_names[["RUT_norm", "Nombre", "Primer Apellido", "Segundo Apellido"]], on="RUT_norm", how="left")

# cumplimiento %
cumpl["Cumplimiento_%"] = (1 - (cumpl["Injustificadas"] / cumpl["Turnos_planificados"].replace({0: pd.NA}))) * 100
cumpl["Cumplimiento_%"] = cumpl["Cumplimiento_%"].round(2)

cumpl = cumpl[[
    "Nombre", "Primer Apellido", "Segundo Apellido",
    "RUT_norm", "Turnos_planificados", "Injustificadas", "Cumplimiento_%"
]].rename(columns={"RUT_norm": "RUT_norm_sin_puntos"}).sort_values(["Cumplimiento_%", "Injustificadas"], ascending=[True, False])

st.dataframe(cumpl, use_container_width=True)

# =========================
# KPIs diarios (matriz: KPIs filas, fechas columnas)
# =========================
st.subheader("KPIs diarios (matriz)")

# Fechas del periodo (día a día)
all_days = pd.date_range(pd.to_datetime(fecha_desde), pd.to_datetime(fecha_hasta), freq="D")
day_labels = [d.strftime("%d-%m-%Y") for d in all_days]

# Turnos planificados diarios (sin libres)
tp_day = (
    df_turnos_valid.groupby(df_turnos_valid["Fecha_dt"].dt.date)
    .size()
)
# Injustificadas diarias
tmp2 = edited.copy()
tmp2["Fecha_dt"] = pd.to_datetime(tmp2["Fecha"], errors="coerce").dt.date
inj_day = (
    tmp2[tmp2["Clasificación Manual"] == "Injustificada"]
    .groupby("Fecha_dt")
    .size()
)

# armar matriz
kpi_rows = ["Turnos_planificados", "Injustificadas", "Cumplimiento_%"]
mat = pd.DataFrame(index=kpi_rows, columns=day_labels)

for d in all_days:
    dd = d.date()
    tp = int(tp_day.get(dd, 0))
    ij = int(inj_day.get(dd, 0))
    mat.loc["Turnos_planificados", d.strftime("%d-%m-%Y")] = tp
    mat.loc["Injustificadas", d.strftime("%d-%m-%Y")] = ij
    if tp > 0:
        mat.loc["Cumplimiento_%", d.strftime("%d-%m-%Y")] = round((1 - (ij / tp)) * 100, 2)
    else:
        mat.loc["Cumplimiento_%", d.strftime("%d-%m-%Y")] = ""

mat = mat.reset_index().rename(columns={"index": "KPI"})
st.dataframe(mat, use_container_width=True)

# =========================
# Export Excel (Cabify + dropdown)
# =========================
st.subheader("Descarga")

# Preparar Incidencias: Fecha como datetime para que Excel la reconozca
edited_export = edited.copy()
edited_export["Fecha"] = pd.to_datetime(edited_export["Fecha"], errors="coerce")

excel_bytes = to_excel_bytes({
    "Incidencias": edited_export,
    "Resumen": resumen,
    "Cumplimiento": cumpl,
    "KPIs_Diarios": mat
}, dropdown_sheet_name="Incidencias")

st.download_button(
    "Descargar Excel consolidado (Cabify + dropdown)",
    data=excel_bytes,
    file_name="reporte_incidencias_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
