import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Incidencias / Ausentismo / Asistencia", layout="wide")

# ------------------------
# Helpers
# ------------------------
def normalize_rut(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    s = s.replace(".", "").replace(" ", "")
    return s  # mantiene guión si existe

def try_parse_date_any(x):
    if pd.isna(x):
        return pd.NaT
    # soporta "29-nov-25", "2025-11-29", datetime, etc.
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def excel_to_df(file, sheet_index=0):
    # Lee por índice de hoja, no por nombre (Hoja1/Hoja2/Hoja 1/Hoja 2 da igual)
    return pd.read_excel(file, sheet_name=sheet_index, engine="openpyxl")

def to_excel_bytes(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    output.seek(0)
    return output

def first_existing_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def get_num(df, col):
    if col not in df.columns:
        return pd.Series([0] * len(df), index=df.index)
    return pd.to_numeric(df[col], errors="coerce").fillna(0)

# ------------------------
# UI
# ------------------------
st.title("App Incidencias / Ausentismo / Asistencia")

with st.sidebar:
    st.header("Cargar archivos (Excel)")
    f_turnos = st.file_uploader("1) Codificación Turnos BUK (Excel)", type=["xlsx"])
    f_activos = st.file_uploader("2) Trabajadores Activos + Turnos (Excel)", type=["xlsx"])

    st.divider()
    st.subheader("Detalle Turnos Colaboradores (1 archivo)")
    st.caption("La Hoja 1 = Inasistencias, la Hoja 2 = Asistencias (por índice).")
    f_detalle = st.file_uploader("3) Detalle Turnos Colaboradores (Excel)", type=["xlsx"])

    st.divider()
    st.subheader("Reglas (MVP)")
    umbral_diff_turno = st.number_input("Umbral Diferencia Turno Real (horas)", value=0.5, step=0.5)
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    st.caption("Déjalo vacío para no filtrar por Área.")

if not all([f_turnos, f_activos, f_detalle]):
    st.info("Sube los 3 archivos para comenzar: Turnos BUK, Activos+Turnos, y Detalle Turnos Colaboradores.")
    st.stop()

# ------------------------
# Load
# ------------------------
df_turnos = excel_to_df(f_turnos, 0)
df_activos = excel_to_df(f_activos, 0)

# Detalle Turnos Colaboradores:
#  - Hoja 1 (índice 0): Inasistencias
#  - Hoja 2 (índice 1): Asistencias
df_inasist = excel_to_df(f_detalle, 0)
df_asist = excel_to_df(f_detalle, 1)

# Normalizar RUT en todas las bases que lo tengan
for df in [df_activos, df_inasist, df_asist]:
    rut_col = first_existing_col(df, ["RUT", "Rut", "rut"])
    if rut_col:
        df["RUT_norm"] = df[rut_col].apply(normalize_rut)
    else:
        df["RUT_norm"] = ""

# ------------------------
# Turnos activos (formato ancho -> largo)
# ------------------------
fixed_cols = [c for c in ["Nombre del Colaborador", "RUT", "Área", "Supervisor"] if c in df_activos.columns]
date_cols = [c for c in df_activos.columns if c not in fixed_cols]

df_act_long = df_activos.melt(
    id_vars=fixed_cols,
    value_vars=date_cols,
    var_name="Fecha",
    value_name="Turno_Activo"
)

df_act_long["Fecha_dt"] = df_act_long["Fecha"].apply(try_parse_date_any)
if "RUT" in df_act_long.columns:
    df_act_long["RUT_norm"] = df_act_long["RUT"].apply(normalize_rut)
else:
    df_act_long["RUT_norm"] = ""

# limpiar blancos
df_act_long["Turno_Activo"] = df_act_long["Turno_Activo"].astype(str).str.strip()
df_act_long.loc[df_act_long["Turno_Activo"].isin(["", "nan", "NaT", "None"]), "Turno_Activo"] = ""

# ------------------------
# Asistencias / Inasistencias: fecha base + clave
# ------------------------
# Asistencias: usar "Fecha Entrada" (tu Excel real lo trae así)
col_fecha_asist = first_existing_col(df_asist, ["Fecha Entrada", "Día", "Dia", "Fecha"])
if col_fecha_asist:
    df_asist["Fecha_base"] = df_asist[col_fecha_asist].apply(try_parse_date_any)
else:
    df_asist["Fecha_base"] = pd.NaT

# Inasistencias: usar "Día" (tu Excel real lo trae así)
col_fecha_inasist = first_existing_col(df_inasist, ["Día", "Dia", "Fecha"])
if col_fecha_inasist:
    df_inasist["Fecha_base"] = df_inasist[col_fecha_inasist].apply(try_parse_date_any)
else:
    df_inasist["Fecha_base"] = pd.NaT

df_asist["Clave_RUT_Fecha_app"] = (
    df_asist["RUT_norm"].astype(str) + "_" +
    df_asist["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")
)

df_inasist["Clave_RUT_Fecha_app"] = (
    df_inasist["RUT_norm"].astype(str) + "_" +
    df_inasist["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")
)

df_act_long["Clave_RUT_Fecha_app"] = (
    df_act_long["RUT_norm"].astype(str) + "_" +
    df_act_long["Fecha_dt"].dt.strftime("%Y-%m-%d").fillna("")
)

# ------------------------
# (Opcional) Filtrar por Área
# ------------------------
def maybe_filter_area(df, col="Área"):
    if only_area and col in df.columns:
        return df[df[col].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()
    return df

df_act_long = maybe_filter_area(df_act_long, "Área")
df_asist = maybe_filter_area(df_asist, "Área")
df_inasist = maybe_filter_area(df_inasist, "Área")

# ------------------------
# Construcción Incidencias por comprobar (MVP)
# ------------------------
inc_rows = []

# 1) Incidencias desde Asistencias (retraso / salida anticipada / diff turno real)
df_asist["Retraso_h"] = get_num(df_asist, "Retraso (horas)")
df_asist["SalidaAnt_h"] = get_num(df_asist, "Salida Anticipada (horas)")
df_asist["DiffTurno_h"] = get_num(df_asist, "Diferencia Turno Real (horas)")

mask_asist = (
    (df_asist["Retraso_h"] > 0) |
    (df_asist["SalidaAnt_h"] > 0) |
    (df_asist["DiffTurno_h"] >= float(umbral_diff_turno))
)
df_asist_inc = df_asist[mask_asist].copy()
df_asist_inc["Fuente"] = "Asistencias PBI (Hoja 2)"
df_asist_inc["Tipo_Incidencia"] = "Marcaje/Turno"
df_asist_inc["Detalle"] = (
    "Retraso_h=" + df_asist_inc["Retraso_h"].astype(str) +
    " | SalidaAnt_h=" + df_asist_inc["SalidaAnt_h"].astype(str) +
    " | DiffTurno_h=" + df_asist_inc["DiffTurno_h"].astype(str)
)

turno_col_asist = first_existing_col(df_asist_inc, ["Turno"])
esp_col_asist = first_existing_col(df_asist_inc, ["Especialidad"])
sup_col_asist = first_existing_col(df_asist_inc, ["Supervisor"])

df_asist_out = pd.DataFrame({
    "Clave_RUT_Fecha_app": df_asist_inc["Clave_RUT_Fecha_app"],
    "RUT": df_asist_inc["RUT_norm"],
    "Fecha": df_asist_inc["Fecha_base"],
    "Turno": df_asist_inc[turno_col_asist] if turno_col_asist else "",
    "Especialidad": df_asist_inc[esp_col_asist] if esp_col_asist else "",
    "Supervisor": df_asist_inc[sup_col_asist] if sup_col_asist else "",
    "Fuente": df_asist_inc["Fuente"],
    "Tipo_Incidencia": df_asist_inc["Tipo_Incidencia"],
    "Detalle": df_asist_inc["Detalle"],
})

inc_rows.append(df_asist_out)

# 2) Incidencias desde Inasistencias (Hoja 1)
df_inasist_inc = df_inasist.copy()
df_inasist_inc["Fuente"] = "Inasistencias PBI (Hoja 1)"
df_inasist_inc["Tipo_Incidencia"] = "Inasistencia"

motivo_col = first_existing_col(df_inasist_inc, ["Motivo"])
df_inasist_inc["Detalle"] = df_inasist_inc[motivo_col].astype(str) if motivo_col else ""

turno_col_in = first_existing_col(df_inasist_inc, ["Turno"])
esp_col_in = first_existing_col(df_inasist_inc, ["Especialidad"])
sup_col_in = first_existing_col(df_inasist_inc, ["Supervisor"])

df_inasist_out = pd.DataFrame({
    "Clave_RUT_Fecha_app": df_inasist_inc["Clave_RUT_Fecha_app"],
    "RUT": df_inasist_inc["RUT_norm"],
    "Fecha": df_inasist_inc["Fecha_base"],
    "Turno": df_inasist_inc[turno_col_in] if turno_col_in else "",
    "Especialidad": df_inasist_inc[esp_col_in] if esp_col_in else "",
    "Supervisor": df_inasist_inc[sup_col_in] if sup_col_in else "",
    "Fuente": df_inasist_inc["Fuente"],
    "Tipo_Incidencia": df_inasist_inc["Tipo_Incidencia"],
    "Detalle": df_inasist_inc["Detalle"],
})

inc_rows.append(df_inasist_out)

df_incidencias = pd.concat(inc_rows, ignore_index=True)

# Join con turnos activos para validar si el turno estaba activo (y mostrarlo)
df_incidencias = df_incidencias.merge(
    df_act_long[["Clave_RUT_Fecha_app", "Turno_Activo"]].rename(columns={"Turno_Activo": "Turno_Activo_Base"}),
    on="Clave_RUT_Fecha_app",
    how="left"
)

# Clasificación manual (editable)
if "Clasificación Manual" not in df_incidencias.columns:
    df_incidencias["Clasificación Manual"] = "Indefinido"

# Orden
cols_order = [
    "Clave_RUT_Fecha_app", "Fecha", "RUT",
    "Turno_Activo_Base", "Turno",
    "Especialidad", "Supervisor",
    "Fuente", "Tipo_Incidencia", "Detalle",
    "Clasificación Manual"
]
df_incidencias = df_incidencias[cols_order].sort_values(["Fecha", "RUT"], na_position="last")

# ------------------------
# UI: Reporte por comprobar + edición
# ------------------------
st.subheader("Reporte Total de Incidencias por Comprobar")

col1, col2 = st.columns([2, 1])
with col2:
    st.write("**Leyenda Clasificación Manual**")
    st.caption("Indefinido (default) / Procede / No procede/Cambio Turno")

edited = st.data_editor(
    df_incidencias,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Clasificación Manual": st.column_config.SelectboxColumn(
            options=["Indefinido", "Procede", "No procede/Cambio Turno"]
        )
    }
)

# ------------------------
# Resumen de incidencias comprobadas
# ------------------------
st.subheader("Reporte Incidencias por tipo (solo comprobadas: Procede)")
df_ok = edited[edited["Clasificación Manual"] == "Procede"].copy()

if len(df_ok) == 0:
    st.warning("Aún no hay incidencias marcadas como 'Procede'.")
    resumen = pd.DataFrame()
else:
    resumen = (
        df_ok.groupby(["Tipo_Incidencia"], dropna=False)
            .size().reset_index(name="Cantidad")
            .sort_values("Cantidad", ascending=False)
    )
    st.dataframe(resumen, use_container_width=True)

# ------------------------
# Export
# ------------------------
st.subheader("Descarga")
excel_bytes = to_excel_bytes({
    "Incidencias_por_comprobar": edited,
    "Resumen_procede": resumen
})
st.download_button(
    "Descargar Excel consolidado",
    data=excel_bytes,
    file_name="reporte_incidencias_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

