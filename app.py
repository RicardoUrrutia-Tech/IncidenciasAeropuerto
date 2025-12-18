import streamlit as st
import pandas as pd

from utils import (
    read_excel,
    build_shift_catalog,
    prepare_activos_turnos,
    prepare_asistencias,
    detect_incidencias,
    build_outputs,
    to_excel_bytes,
)

st.set_page_config(page_title="Incidencias / Ausentismo / Cumplimiento", layout="wide")

st.title("APP Reportes Incidencias / Ausentismo / Cumplimiento")

with st.sidebar:
    st.header("📥 Carga de archivos (Excel)")
    f_cod = st.file_uploader("1) Base Codificación Turnos (BUK)", type=["xlsx"])
    f_act = st.file_uploader("2) Trabajadores Activos + Turnos", type=["xlsx"])
    f_det = st.file_uploader("3) Detalle Turnos Colaboradores (BUK)", type=["xlsx"])
    f_asi = st.file_uploader("4) Base Asistencias PBI (marcajes)", type=["xlsx"])

    st.divider()
    st.caption("Opcional")
    f_manual = st.file_uploader("Ingreso manual incidencias (opcional)", type=["xlsx"])

    st.divider()
    tolerance_min = st.number_input("Tolerancia (min) para atrasos/anticipos", min_value=0, max_value=120, value=5, step=1)

if not (f_cod and f_act and f_det and f_asi):
    st.info("Carga los 4 archivos obligatorios para comenzar.")
    st.stop()

# ---- Load
df_cod = read_excel(f_cod)
df_act = read_excel(f_act)
df_det = read_excel(f_det)
df_asi = read_excel(f_asi)

shift_catalog = build_shift_catalog(df_cod)

# ---- Prepare
act_long = prepare_activos_turnos(df_act, shift_catalog)
asist = prepare_asistencias(df_asi, shift_catalog)

# (Opcional) manual
manual_df = None
if f_manual is not None:
    manual_df = read_excel(f_manual)

# ---- Incidencias
incidencias = detect_incidencias(
    act_long=act_long,
    asist=asist,
    df_det=df_det,
    tolerance_min=int(tolerance_min),
    manual_df=manual_df,
)

# ---- UI: Reporte principal (por comprobar)
st.subheader("1) Reporte Total de Incidencias por Comprobar")
st.caption("Columna 'Comprobación Incidencia' por defecto en 'Indefinido'.")

# Editable table (guardado en session_state)
if "incidencias_edit" not in st.session_state:
    st.session_state["incidencias_edit"] = incidencias.copy()

edit_df = st.data_editor(
    st.session_state["incidencias_edit"],
    use_container_width=True,
    num_rows="fixed",
    column_config={
        "Comprobación Incidencia": st.column_config.SelectboxColumn(
            "Comprobación Incidencia",
            options=["Indefinido", "Procede", "No procede/Cambio turno"],
            required=True,
        )
    },
)

st.session_state["incidencias_edit"] = edit_df

c1, c2 = st.columns(2)
with c1:
    st.download_button(
        "⬇️ Descargar Incidencias (Excel)",
        data=to_excel_bytes(edit_df, sheet_name="Incidencias_por_comprobar"),
        file_name="incidencias_por_comprobar.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with c2:
    st.download_button(
        "⬇️ Descargar catálogo turnos normalizado",
        data=to_excel_bytes(shift_catalog, sheet_name="CatalogoTurnos"),
        file_name="catalogo_turnos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()

# ---- Outputs finales
outputs = build_outputs(act_long, asist, edit_df)

tabs = st.tabs([
    "2) Incidencias por tipo (comprobadas)",
    "3) Marcaje",
    "4) Cumplimiento por trabajador",
    "5) Ausentismo por trabajador",
    "6) Asistencia diaria",
])

with tabs[0]:
    st.subheader("2) Incidencias por tipo (solo 'Procede')")
    st.dataframe(outputs["incidencias_por_tipo"], use_container_width=True)

with tabs[1]:
    st.subheader("3) Reporte de Marcaje")
    st.dataframe(outputs["reporte_marcaje"], use_container_width=True)

with tabs[2]:
    st.subheader("4) Cumplimiento por Trabajador")
    st.dataframe(outputs["cumplimiento_trabajador"], use_container_width=True)

with tabs[3]:
    st.subheader("5) Ausentismo por Trabajador")
    st.dataframe(outputs["ausentismo_trabajador"], use_container_width=True)

with tabs[4]:
    st.subheader("6) Asistencia diaria (%), por Especialidad y General")
    st.dataframe(outputs["asistencia_diaria"], use_container_width=True)

st.divider()

# Export pack
st.subheader("📦 Descargar pack completo de reportes (Excel)")
pack_bytes = to_excel_bytes(
    outputs,
    multi_sheet=True,
    filename_hint="pack_reportes.xlsx",
)
st.download_button(
    "⬇️ Descargar pack",
    data=pack_bytes,
    file_name="pack_reportes.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
