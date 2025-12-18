import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Incidencias / Ausentismo / Asistencia", layout="wide")

# ------------------------
# Helpers
# ------------------------
def normalize_rut(x: str) -> str:
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
    # Lee por índice de hoja, no por nombre (Hoja 1 / Hoja 2 da igual)
    return pd.read_excel(file, sheet_name=sheet_index, engine="openpyxl")

def to_excel_bytes(df_dict):
    # df_dict: {"NombreHoja": df, ...}
    output = BytesIO()
    # Usamos openpyxl para no requerir xlsxwriter
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in df_dict.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    output.seek(0)
    return output

def get_num(df, col):
    if col not in df.columns:
        return pd.Series([0] * len(df), index=df.index)
    return pd.to_numeric(df[col], errors="coerce").fillna(0)

def safe_col(df, candidates, default=""):
    """
    Devuelve el nombre de la primera columna existente en df dentro de candidates.
    Si ninguna existe, devuelve None.
    """
    for c in candidates:
        if c in df.columns:
            return c
    return None

# ------------------------
# UI
# ------------------------
st.title("App Incidencias / Ausentismo / Asistencia")

with st.sidebar:
    st.header("Cargar archivos (Excel)")

    f_turnos = st.file_uploader("1) Codificación Turnos BUK", type=["xlsx"])
    f_activos = st.file_uploader("2) Trabajadores Activos + Turnos", type=["xlsx"])

    # ✅ Un solo archivo PBI con 2 hojas:
    # Hoja 1 (índice 0): Inasistencias
    # Hoja 2 (índice 1): Asistencias
    f_pbi = st.file_uploader("3) Reporte PBI (Hoja 1=Inasistencias, Hoja 2=Asistencias)", type=["xlsx"])

    st.divider()
    st.subheader("Reglas (MVP)")
    umbral_diff_turno = st.number_input("Umbral Diferencia Turno Real (horas)", value=0.5, step=0.5)
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    st.caption("El filtro de Área es opcional. Déjalo vacío para no filtrar.")

if not all([f_turnos, f_activos, f_pbi]):
    st.info("Sube los 3 archivos para comenzar (Turnos, Activos+Turnos, Reporte PBI con 2 hojas).")
    st.stop()

# ------------------------
# Load
# ------------------------
df_turnos = excel_to_df(f_turnos, 0)     # (por ahora no se usa en reglas MVP, pero queda cargado)
df_activos = excel_to_df(f_activos, 0)

# ✅ Del mismo Excel PBI:
df_inasist = excel_to_df(f_pbi, 0)  # Hoja 1 = Inasistencias
df_asist = excel_to_df(f_pbi, 1)    # Hoja 2 = Asistencias

# Normalize basic fields
for df in [df_activos, df_inasist, df_asist]:
    if "RUT" in df.columns:
        df["RUT_norm"] = df["RUT"].apply(normalize_rut)
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
df_act_long["RUT_norm"] = df_act_long["RUT"].apply(normalize_rut) if "RUT" in df_act_long.columns else ""

# Limpiar blancos
df_act_long["Turno_Activo"] = df_act_long["Turno_Activo"].astype(str).str.strip()
df_act_long.loc[df_act_long["Turno_Activo"].isin(["", "nan", "NaT", "None"]), "Turno_Activo"] = ""

# ------------------------
# Asistencias / Inasistencias: clave rut-fecha
# ------------------------
# Asistencias: usar Fecha Entrada como fecha base (o alternativas)
if "Fecha Entrada" in df_asist.columns:
    df_asist["Fecha_base"] = df_asist["Fecha Entrada"].apply(try_parse_date_any)
elif "Día" in df_asist.columns:
    df_asist["Fecha_base"] = df_asist["Día"].apply(try_parse_date_any)
else:
    df_asist["Fecha_base"] = pd.NaT

# Inasistencias: usar Día (o alternativas)
if "Día" in df_inasist.columns:
    df_inasist["Fecha_base"] = df_inasist["Día"].apply(try_parse_date_any)
elif "Fecha" in df_inasist.columns:
    df_inasist["Fecha_base"] = df_inasist["Fecha"].apply(try_parse_date_any)
else:
    df_inasist["Fecha_base"] = pd.NaT

df_asist["Clave_RUT_Fecha_app"] = df_asist["RUT_norm"].astype(str) + "_" + df_asist["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")
df_inasist["Clave_RUT_Fecha_app"] = df_inasist["RUT_norm"].astype(str) + "_" + df_inasist["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")
df_act_long["Clave_RUT_Fecha_app"] = df_act_long["RUT_norm"].astype(str) + "_" + df_act_long["Fecha_dt"].dt.strftime("%Y-%m-%d").fillna("")

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

# 1) Incidencias desde Asistencias (retrazo / salida anticipada / diff turno real)
df_asist["Retraso_h"] = get_num(df_asist, "Retraso (horas)")
df_asist["SalidaAnt_h"] = get_num(df_asist, "Salida Anticipada (horas)")
df_asist["DiffTurno_h"] = get_num(df_asist, "Diferencia Turno Real (horas)")

mask_asist = (
    (df_asist["Retraso_h"] > 0) |
    (df_asist["SalidaAnt_h"] > 0) |
    (df_asist["DiffTurno_h"] >= float(umbral_diff_turno))
)

df_asist_inc = df_asist[mask_asist].copy()
df_asist_inc["Fuente"] = "Asistencias PBI"
df_asist_inc["Tipo_Incidencia"] = "Marcaje/Turno"
df_asist_inc["Detalle"] = (
    "Retraso_h=" + df_asist_inc["Retraso_h"].astype(str) +
    " | SalidaAnt_h=" + df_asist_inc["SalidaAnt_h"].astype(str) +
    " | DiffTurno_h=" + df_asist_inc["DiffTurno_h"].astype(str)
)

turno_col_asist = safe_col(df_asist_inc, ["Turno", "Turno Programado", "Horario"])
esp_col_asist = safe_col(df_asist_inc, ["Especialidad", "Cargo", "Puesto"])
sup_col_asist = safe_col(df_asist_inc, ["Supervisor", "Jefatura", "Líder"])

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

# 2) Incidencias desde Inasistencias PBI
df_inasist_inc = df_inasist.copy()
df_inasist_inc["Fuente"] = "Inasistencias PBI"
df_inasist_inc["Tipo_Incidencia"] = "Inasistencia"

motivo_col = safe_col(df_inasist_inc, ["Motivo", "Justificación", "Causal", "Observación"])
df_inasist_inc["Detalle"] = df_inasist_inc[motivo_col].astype(str) if motivo_col else ""

turno_col_inas = safe_col(df_inasist_inc, ["Turno", "Turno Programado", "Horario"])
esp_col_inas = safe_col(df_inasist_inc, ["Especialidad", "Cargo", "Puesto"])
sup_col_inas = safe_col(df_inasist_inc, ["Supervisor", "Jefatura", "Líder"])

df_inas_out = pd.DataFrame({
    "Clave_RUT_Fecha_app": df_inasist_inc["Clave_RUT_Fecha_app"],
    "RUT": df_inasist_inc["RUT_norm"],
    "Fecha": df_inasist_inc["Fecha_base"],
    "Turno": df_inasist_inc[turno_col_inas] if turno_col_inas else "",
    "Especialidad": df_inasist_inc[esp_col_inas] if esp_col_inas else "",
    "Supervisor": df_inasist_inc[sup_col_inas] if sup_col_inas else "",
    "Fuente": df_inasist_inc["Fuente"],
    "Tipo_Incidencia": df_inasist_inc["Tipo_Incidencia"],
    "Detalle": df_inasist_inc["Detalle"],
})

inc_rows.append(df_inas_out)

df_incidencias = pd.concat(inc_rows, ignore_index=True)

# Join con turnos activos para validar si el turno estaba activo (y mostrarlo)
df_incidencias = df_incidencias.merge(
    df_act_long[["Clave_RUT_Fecha_app", "Turno_Activo"]].rename(columns={"Turno_Activo": "Turno_Activo_Base"}),
    on="Clave_RUT_Fecha_app",
    how="left"
)

# Columna clasificación manual
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
df_incidencias = df_incidencias[[c for c in cols_order if c in df_incidencias.columns]].copy()
df_incidencias = df_incidencias.sort_values(["Fecha", "RUT"], na_position="last")

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

resumen = pd.DataFrame()
if len(df_ok) == 0:
    st.warning("Aún no hay incidencias marcadas como 'Procede'.")
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

