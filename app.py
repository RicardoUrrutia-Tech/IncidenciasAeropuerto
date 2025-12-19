import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# =========================================================
# Config
# =========================================================
st.set_page_config(page_title="Incidencias / Ausentismo / Asistencia", layout="wide")

CABIFY_DARK = "1F123F"   # Gradiente moradul (oscuro)
CABIFY_SOFT = "F5F1FC"   # Gradiente (muy claro)
CABIFY_SOFT2 = "FAF8FE"  # Gradiente (aún más claro)

CLASIF_OPTS = ["Indefinido", "Procede", "No procede/Cambio Turno"]


# =========================================================
# Helpers
# =========================================================
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

def get_num(df, col):
    if col not in df.columns:
        return pd.Series([0] * len(df))
    return pd.to_numeric(df[col], errors="coerce").fillna(0)

def maybe_filter_area(df, only_area, col="Área"):
    if only_area and col in df.columns:
        return df[df[col].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()
    return df

def safe_col(df, name, fallback=""):
    # Devuelve la columna si existe; si no, una serie vacía
    if name in df.columns:
        return df[name]
    return pd.Series([fallback] * len(df))

def build_excel_bytes(df_incidencias, df_resumen):
    """
    Exporta a Excel con:
    - Estilo Cabify
    - Fecha formato corto
    - Dropdown (data validation) en 'Clasificación Manual'
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Incidencias_por_comprobar"

    # Orden de columnas (sin Clave_RUT_Fecha_app ni Turno_Activo_Base)
    wanted = [
        "Fecha", "RUT", "Turno", "Especialidad", "Supervisor",
        "Fuente", "Tipo_Incidencia", "Detalle", "Clasificación Manual"
    ]
    cols = [c for c in wanted if c in df_incidencias.columns]
    df_out = df_incidencias[cols].copy()

    # Asegurar Fecha como datetime para formateo en Excel
    if "Fecha" in df_out.columns:
        df_out["Fecha"] = pd.to_datetime(df_out["Fecha"], errors="coerce")

    # Escribir header
    header_fill = PatternFill("solid", fgColor=CABIFY_DARK)
    header_font = Font(color="FFFFFF", bold=True)
    center = Alignment(vertical="center", wrap_text=True)
    left = Alignment(vertical="center", wrap_text=True)

    ws.append(cols)
    for j, colname in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=j)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center

    # Escribir data
    for _, row in df_out.iterrows():
        ws.append(row.tolist())

    # Freeze + autofilter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"

    # Zebra rows + formatos
    zebra1 = PatternFill("solid", fgColor=CABIFY_SOFT)
    zebra2 = PatternFill("solid", fgColor=CABIFY_SOFT2)

    fecha_col_idx = cols.index("Fecha") + 1 if "Fecha" in cols else None
    clasif_col_idx = cols.index("Clasificación Manual") + 1 if "Clasificación Manual" in cols else None

    for r in range(2, ws.max_row + 1):
        fill = zebra1 if (r % 2 == 0) else zebra2
        for c in range(1, len(cols) + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill
            cell.alignment = left

        # Fecha corta
        if fecha_col_idx is not None:
            ws.cell(row=r, column=fecha_col_idx).number_format = "dd-mm-yyyy"

    # Ajuste de anchos (simple y efectivo)
    for c in range(1, len(cols) + 1):
        letter = get_column_letter(c)
        max_len = 0
        for r in range(1, min(ws.max_row, 300) + 1):  # limita para performance
            v = ws.cell(row=r, column=c).value
            if v is None:
                continue
            v = str(v)
            if len(v) > max_len:
                max_len = len(v)
        ws.column_dimensions[letter].width = min(max(12, max_len + 2), 55)

    # Data validation (dropdown) para Clasificación Manual
    if clasif_col_idx is not None and ws.max_row >= 2:
        dv = DataValidation(
            type="list",
            formula1=f'"{",".join(CLASIF_OPTS)}"',
            allow_blank=False,
            showDropDown=True
        )
        ws.add_data_validation(dv)
        dv_range = f"{get_column_letter(clasif_col_idx)}2:{get_column_letter(clasif_col_idx)}{ws.max_row}"
        dv.add(dv_range)

    # Hoja resumen
    ws2 = wb.create_sheet("Resumen_procede")
    if df_resumen is None or df_resumen.empty:
        ws2.append(["Sin datos"])
    else:
        ws2.append(list(df_resumen.columns))
        for j in range(1, len(df_resumen.columns) + 1):
            cell = ws2.cell(row=1, column=j)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center

        for _, row in df_resumen.iterrows():
            ws2.append(row.tolist())

        ws2.freeze_panes = "A2"
        ws2.auto_filter.ref = f"A1:{get_column_letter(len(df_resumen.columns))}1"

        for r in range(2, ws2.max_row + 1):
            fill = zebra1 if (r % 2 == 0) else zebra2
            for c in range(1, len(df_resumen.columns) + 1):
                cell = ws2.cell(row=r, column=c)
                cell.fill = fill
                cell.alignment = left

        for c in range(1, len(df_resumen.columns) + 1):
            letter = get_column_letter(c)
            ws2.column_dimensions[letter].width = 22

    # Bytes
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# =========================================================
# UI
# =========================================================
st.title("App Incidencias / Ausentismo / Asistencia")

with st.sidebar:
    st.header("Cargar archivos (Excel)")
    f_turnos = st.file_uploader("1) Codificación Turnos BUK", type=["xlsx"])
    f_activos = st.file_uploader("2) Trabajadores Activos + Turnos", type=["xlsx"])
    f_detalle = st.file_uploader("3) Detalle Turnos Colaboradores (Hoja1=Inasistencias, Hoja2=Asistencias)", type=["xlsx"])

    st.divider()
    st.subheader("Reglas (MVP)")
    umbral_diff_turno = st.number_input("Umbral Diferencia Turno Real (horas)", value=0.5, step=0.5)
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    st.caption("Déjalo vacío para no filtrar.")

if not all([f_turnos, f_activos, f_detalle]):
    st.info("Sube los 3 archivos para comenzar.")
    st.stop()


# =========================================================
# Load
# =========================================================
df_turnos = excel_to_df(f_turnos, 0)     # (por ahora no lo usamos en el MVP, pero lo dejamos cargado)
df_activos = excel_to_df(f_activos, 0)

# Detalle Turnos Colaboradores:
# Hoja 1 (index 0) = Inasistencias
# Hoja 2 (index 1) = Asistencias
df_inasist = excel_to_df(f_detalle, 0)
df_asist = excel_to_df(f_detalle, 1)

# Normalizar RUT si existe
for df in [df_activos, df_inasist, df_asist]:
    if "RUT" in df.columns:
        df["RUT_norm"] = df["RUT"].apply(normalize_rut)
    else:
        df["RUT_norm"] = ""


# =========================================================
# Turnos activos (formato ancho -> largo)
# =========================================================
fixed_cols = [c for c in ["Nombre del Colaborador", "RUT", "Área", "Supervisor"] if c in df_activos.columns]
date_cols = [c for c in df_activos.columns if c not in fixed_cols]

df_act_long = df_activos.melt(
    id_vars=fixed_cols,
    value_vars=date_cols,
    var_name="Fecha",
    value_name="Turno_Activo"
)
df_act_long["Fecha_dt"] = df_act_long["Fecha"].apply(try_parse_date_any)
df_act_long["RUT_norm"] = safe_col(df_act_long, "RUT").apply(normalize_rut)

# Limpiar blancos
df_act_long["Turno_Activo"] = df_act_long["Turno_Activo"].astype(str).str.strip()
df_act_long.loc[df_act_long["Turno_Activo"].isin(["", "nan", "NaT", "None", "-"]), "Turno_Activo"] = ""

# Clave para join interno (no se mostrará)
df_act_long["Clave_RUT_Fecha_app"] = df_act_long["RUT_norm"].astype(str) + "_" + df_act_long["Fecha_dt"].dt.strftime("%Y-%m-%d").fillna("")


# =========================================================
# Asistencias / Inasistencias: fechas base + clave (no se mostrará)
# =========================================================
# Asistencias: usar Fecha Entrada (según tu archivo real)
df_asist["Fecha_base"] = safe_col(df_asist, "Fecha Entrada").apply(try_parse_date_any)
df_asist["Clave_RUT_Fecha_app"] = df_asist["RUT_norm"].astype(str) + "_" + df_asist["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")

# Inasistencias: usar Día (según tu archivo real)
df_inasist["Fecha_base"] = safe_col(df_inasist, "Día").apply(try_parse_date_any)
df_inasist["Clave_RUT_Fecha_app"] = df_inasist["RUT_norm"].astype(str) + "_" + df_inasist["Fecha_base"].dt.strftime("%Y-%m-%d").fillna("")


# =========================================================
# Filtrar Área
# =========================================================
df_act_long = maybe_filter_area(df_act_long, only_area, "Área")
df_asist = maybe_filter_area(df_asist, only_area, "Área")
df_inasist = maybe_filter_area(df_inasist, only_area, "Área")


# =========================================================
# Construcción Incidencias por comprobar (MVP)
# =========================================================
inc_rows = []

# 1) Incidencias desde Asistencias (Retraso / Salida anticipada / Diff turno real)
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

inc_rows.append(
    df_asist_inc.assign(
        RUT=df_asist_inc["RUT_norm"],
        Fecha=df_asist_inc["Fecha_base"],
        Turno=safe_col(df_asist_inc, "Turno", ""),
        Especialidad=safe_col(df_asist_inc, "Especialidad", ""),
        Supervisor=safe_col(df_asist_inc, "Supervisor", "")
    )[["Clave_RUT_Fecha_app", "Fecha", "RUT", "Turno", "Especialidad", "Supervisor", "Fuente", "Tipo_Incidencia", "Detalle"]]
)

# 2) Incidencias desde Inasistencias (todas, para clasificar manualmente)
df_inasist_inc = df_inasist.copy()
df_inasist_inc["Fuente"] = "Inasistencias PBI"
df_inasist_inc["Tipo_Incidencia"] = "Inasistencia"
df_inasist_inc["Detalle"] = safe_col(df_inasist_inc, "Motivo", "").astype(str)

inc_rows.append(
    df_inasist_inc.assign(
        RUT=df_inasist_inc["RUT_norm"],
        Fecha=df_inasist_inc["Fecha_base"],
        Turno=safe_col(df_inasist_inc, "Turno", ""),
        Especialidad=safe_col(df_inasist_inc, "Especialidad", ""),
        Supervisor=safe_col(df_inasist_inc, "Supervisor", "")
    )[["Clave_RUT_Fecha_app", "Fecha", "RUT", "Turno", "Especialidad", "Supervisor", "Fuente", "Tipo_Incidencia", "Detalle"]]
)

df_incidencias = pd.concat(inc_rows, ignore_index=True)

# Join con turnos activos (solo para control interno; NO lo mostraremos)
df_incidencias = df_incidencias.merge(
    df_act_long[["Clave_RUT_Fecha_app", "Turno_Activo"]].rename(columns={"Turno_Activo": "Turno_Activo_Base"}),
    on="Clave_RUT_Fecha_app",
    how="left"
)

# Clasificación manual
if "Clasificación Manual" not in df_incidencias.columns:
    df_incidencias["Clasificación Manual"] = "Indefinido"

# Limpieza final: Fecha datetime
df_incidencias["Fecha"] = pd.to_datetime(df_incidencias["Fecha"], errors="coerce")

# Orden
df_incidencias = df_incidencias.sort_values(["Fecha", "RUT"], na_position="last")


# =========================================================
# UI Tabla (sin Clave_RUT_Fecha_app ni Turno_Activo_Base)
# =========================================================
st.subheader("Reporte Total de Incidencias por Comprobar")

cols_show = [
    "Fecha", "RUT", "Turno", "Especialidad", "Supervisor",
    "Fuente", "Tipo_Incidencia", "Detalle", "Clasificación Manual"
]
cols_show = [c for c in cols_show if c in df_incidencias.columns]

col1, col2 = st.columns([2, 1])
with col2:
    st.write("**Leyenda Clasificación Manual**")
    st.caption("Indefinido (default) / Procede / No procede/Cambio Turno")

edited = st.data_editor(
    df_incidencias[cols_show],
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Fecha": st.column_config.DateColumn(format="DD-MM-YYYY"),
        "Clasificación Manual": st.column_config.SelectboxColumn(
            options=CLASIF_OPTS
        )
    }
)

# =========================================================
# Resumen (solo Procede)
# =========================================================
st.subheader("Reporte Incidencias por tipo (solo comprobadas: Procede)")
df_ok = edited[edited["Clasificación Manual"] == "Procede"].copy()

if df_ok.empty:
    st.warning("Aún no hay incidencias marcadas como 'Procede'.")
    resumen = pd.DataFrame(columns=["Tipo_Incidencia", "Cantidad"])
else:
    resumen = (
        df_ok.groupby(["Tipo_Incidencia"], dropna=False)
            .size().reset_index(name="Cantidad")
            .sort_values("Cantidad", ascending=False)
    )
    st.dataframe(resumen, use_container_width=True)


# =========================================================
# Export (openpyxl + estilo + dropdown)
# =========================================================
st.subheader("Descarga (Excel con estilo Cabify + dropdown)")

excel_bytes = build_excel_bytes(
    df_incidencias=edited.copy(),
    df_resumen=resumen.copy()
)

st.download_button(
    "Descargar Excel consolidado",
    data=excel_bytes,
    file_name="reporte_incidencias_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

