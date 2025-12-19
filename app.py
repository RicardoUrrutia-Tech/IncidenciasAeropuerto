import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# =========================
# Config
# =========================
st.set_page_config(page_title="Incidencias / Ausentismo / Asistencia", layout="wide")
st.title("App Incidencias / Ausentismo / Asistencia")

CABIFY_PURPLE = "1F123F"   # Gradiente Moradul (principal)
CABIFY_PINK = "E83C96"     # Contraste
CABIFY_LIGHT = "FAF8FE"    # Fondo muy claro

CLASIF_OPTIONS = ["Seleccionar", "Injustificada", "Permiso", "No Procede - Cambio de Turno"]


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
    if pd.isna(x) or str(x).strip() == "":
        return pd.NaT
    # Soporta 29-nov-25, 2025-11-29, datetime, etc.
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def read_excel_sheet_by_index(file, sheet_index: int) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name=sheet_index, engine="openpyxl")

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    # Normaliza nombres de columnas (quita espacios dobles, etc.)
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def pick_col(df: pd.DataFrame, candidates):
    """Devuelve el primer nombre de columna existente en df para una lista de candidatos."""
    for c in candidates:
        if c in df.columns:
            return c
    return None

def to_num_series(df: pd.DataFrame, colname: str) -> pd.Series:
    if not colname or colname not in df.columns:
        return pd.Series([0] * len(df), index=df.index)
    return pd.to_numeric(df[colname], errors="coerce").fillna(0)

def autosize_columns(ws, df, min_w=10, max_w=45):
    for i, col in enumerate(df.columns, start=1):
        max_len = len(str(col))
        # mirar algunas filas para ancho razonable
        sample = df[col].astype(str).head(400)
        max_len = max(max_len, sample.map(len).max() if len(sample) else max_len)
        width = max(min_w, min(max_w, max_len + 2))
        ws.column_dimensions[get_column_letter(i)].width = width

def build_excel_with_style_and_dropdown(df: pd.DataFrame, sheet_name="Incidencias"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]

    # Header style
    header_fill = PatternFill("solid", fgColor=CABIFY_PURPLE)
    header_font = Font(color="FFFFFF", bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Body style
    body_align = Alignment(vertical="top", wrap_text=True)

    # Write header
    for j, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=j, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align

    # Write data
    for i, row in enumerate(df.itertuples(index=False), start=2):
        for j, val in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=val)

    # Freeze header + autofilter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{max(1, len(df) + 1)}"

    # Apply alignment
    for row in ws.iter_rows(min_row=2, max_row=len(df) + 1, min_col=1, max_col=len(df.columns)):
        for cell in row:
            cell.alignment = body_align

    # Date format (Fecha)
    if "Fecha" in df.columns:
        fecha_idx = list(df.columns).index("Fecha") + 1
        for r in range(2, len(df) + 2):
            c = ws.cell(row=r, column=fecha_idx)
            c.number_format = "dd-mm-yyyy"

    # Dropdown (Data Validation) for "Clasificación Manual"
    if "Clasificación Manual" in df.columns:
        clasif_idx = list(df.columns).index("Clasificación Manual") + 1
        dv = DataValidation(
            type="list",
            formula1=f'"{",".join(CLASIF_OPTIONS)}"',
            allow_blank=False,
            showDropDown=True
        )
        ws.add_data_validation(dv)

        # aplicar a un rango (todas las filas con datos + un extra por si agregan algunas)
        max_rows = max(len(df) + 1, 5000)  # rango amplio para facilitar trabajo del coordinador
        dv_range = f"{get_column_letter(clasif_idx)}2:{get_column_letter(clasif_idx)}{max_rows}"
        dv.add(dv_range)

        # Asegurar valor por defecto "Seleccionar" en filas existentes
        for r in range(2, len(df) + 2):
            cell = ws.cell(row=r, column=clasif_idx)
            if cell.value is None or str(cell.value).strip() == "":
                cell.value = "Seleccionar"

    # Cabify subtle row shading (optional light fill for readability)
    light_fill = PatternFill("solid", fgColor=CABIFY_LIGHT)
    for r in range(2, len(df) + 2):
        if r % 2 == 0:
            for c in range(1, len(df.columns) + 1):
                ws.cell(row=r, column=c).fill = light_fill

    autosize_columns(ws, df)

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# =========================
# UI inputs
# =========================
with st.sidebar:
    st.header("Cargar archivos (Excel)")

    f_turnos = st.file_uploader("1) Codificación Turnos BUK", type=["xlsx"])
    f_activos = st.file_uploader("2) Trabajadores Activos + Turnos", type=["xlsx"])
    f_pbi = st.file_uploader("3) Detalle Turnos Colaboradores (PBI) - Hoja1 Inasist / Hoja2 Asist", type=["xlsx"])

    st.divider()
    st.subheader("Reglas")
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    st.caption("Déjalo vacío para no filtrar.")

if not all([f_turnos, f_activos, f_pbi]):
    st.info("Sube los 3 archivos para comenzar.")
    st.stop()


# =========================
# Load files
# =========================
df_turnos = clean_columns(read_excel_sheet_by_index(f_turnos, 0))
df_activos = clean_columns(read_excel_sheet_by_index(f_activos, 0))

# PBI: Hoja 1 inasistencias, Hoja 2 asistencias (según tu doc)
df_inasist = clean_columns(read_excel_sheet_by_index(f_pbi, 0))
df_asist = clean_columns(read_excel_sheet_by_index(f_pbi, 1))

# Normalizar RUT en asistencias (clave del error que tenías: no existía RUT_norm si no encontraba "RUT")
rut_col_asist = pick_col(df_asist, ["RUT", "Rut", "rut"])
if rut_col_asist:
    df_asist["RUT_norm"] = df_asist[rut_col_asist].apply(normalize_rut)
else:
    # si no existe, dejamos vacío para evitar KeyError y mostrar aviso
    df_asist["RUT_norm"] = ""

# Filtrar área opcional
def maybe_filter_area(df: pd.DataFrame):
    if not only_area:
        return df
    area_col = pick_col(df, ["Área", "Area", "AREA"])
    if not area_col:
        return df
    return df[df[area_col].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()

df_asist = maybe_filter_area(df_asist)


# =========================
# Construir incidencias SOLO desde Asistencias
# Regla: SOLO si existe Retraso o Salida Anticipada (no usar DiffTurno)
# =========================
# Detectar columnas posibles
col_retraso = pick_col(df_asist, ["Retraso (horas)", "Retraso", "Horas Retraso", "Retraso_h"])
col_salida = pick_col(df_asist, ["Salida Anticipada (horas)", "Salida Anticipada", "Horas Salida Anticipada", "SalidaAnt_h"])

df_asist["Retraso_h"] = to_num_series(df_asist, col_retraso)
df_asist["SalidaAnt_h"] = to_num_series(df_asist, col_salida)

mask_inc = (df_asist["Retraso_h"] > 0) | (df_asist["SalidaAnt_h"] > 0)
df_inc = df_asist[mask_inc].copy()

# Fecha base: por especificación del reporte Asistencias, usar "Fecha Entrada" (o "Día" fallback)
col_fecha = pick_col(df_inc, ["Fecha Entrada", "Fecha", "Día", "Dia"])
df_inc["Fecha"] = df_inc[col_fecha].apply(try_parse_date_any) if col_fecha else pd.NaT

# Campos solicitados por ti
col_nombre = pick_col(df_inc, ["Nombre", "Nombres", "Nombre del Colaborador"])
col_ap1 = pick_col(df_inc, ["Primer Apellido", "Apellido Paterno", "Apellido1"])
col_ap2 = pick_col(df_inc, ["Segundo Apellido", "Apellido Materno", "Apellido2"])
col_rut = rut_col_asist
col_turno = pick_col(df_inc, ["Turno", "Horario", "Turno (horario)"])
col_esp = pick_col(df_inc, ["Especialidad", "Cargo"])
col_sup = pick_col(df_inc, ["Supervisor", "Jefatura", "Líder", "Lider"])

# Tipo_Incidencia + Detalle
def tipo_incidencia(row):
    r = float(row.get("Retraso_h", 0))
    s = float(row.get("SalidaAnt_h", 0))
    if r > 0 and s > 0:
        return "Retraso + Salida Anticipada"
    if r > 0:
        return "Retraso"
    return "Salida Anticipada"

df_inc["Tipo_Incidencia"] = df_inc.apply(tipo_incidencia, axis=1)

df_inc["Detalle"] = (
    "Retraso_h=" + df_inc["Retraso_h"].round(2).astype(str) +
    " | SalidaAnt_h=" + df_inc["SalidaAnt_h"].round(2).astype(str)
)

# Clasificación Manual por defecto
df_inc["Clasificación Manual"] = "Seleccionar"

# Armar dataframe final SOLO con campos requeridos
out = pd.DataFrame({
    "Fecha": df_inc["Fecha"].dt.date,  # para “fecha corta” en pantalla; en excel damos formato también
    "Nombre": df_inc[col_nombre] if col_nombre else "",
    "Primer Apellido": df_inc[col_ap1] if col_ap1 else "",
    "Segundo Apellido": df_inc[col_ap2] if col_ap2 else "",
    "RUT": df_inc[col_rut].astype(str) if col_rut else df_inc["RUT_norm"],
    "Turno": df_inc[col_turno] if col_turno else "",
    "Especialidad": df_inc[col_esp] if col_esp else "",
    "Supervisor": df_inc[col_sup] if col_sup else "",
    "Tipo_Incidencia": df_inc["Tipo_Incidencia"],
    "Detalle": df_inc["Detalle"],
    "Clasificación Manual": df_inc["Clasificación Manual"],
})

# =========================
# UI: tabla editable
# =========================
st.subheader("Reporte Total de Incidencias por Comprobar (solo Retraso / Salida Anticipada)")

if out.empty:
    st.warning("No se encontraron incidencias con Retraso o Salida Anticipada en la Hoja 2 (Asistencias).")
    st.stop()

edited = st.data_editor(
    out,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Fecha": st.column_config.DateColumn(format="DD-MM-YYYY"),
        "Clasificación Manual": st.column_config.SelectboxColumn(
            options=CLASIF_OPTIONS,
            required=True
        ),
    }
)

# =========================
# Export Excel (Cabify + dropdown)
# =========================
st.subheader("Descarga")

excel_bytes = build_excel_with_style_and_dropdown(
    edited,
    sheet_name="Incidencias"
)

st.download_button(
    "Descargar Excel consolidado (Cabify + selector)",
    data=excel_bytes,
    file_name="reporte_incidencias_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

