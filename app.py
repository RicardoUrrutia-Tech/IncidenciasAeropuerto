import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation


st.set_page_config(page_title="Incidencias / Ausentismo / Asistencia", layout="wide")

# ========================
# Helpers
# ========================
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
        return pd.Series([0] * len(df), index=df.index)
    return pd.to_numeric(df[col], errors="coerce").fillna(0)

def safe_col(df, colname):
    """Devuelve la columna exacta si existe (match case-insensitive)."""
    cols = {c.strip().lower(): c for c in df.columns}
    return cols.get(colname.strip().lower())

# ========================
# Excel Export (Cabify + Dropdown)
# ========================
CABIFY = {
    "moradul_dark": "1F123F",
    "moradul_mid": "4A2B8D",
    "moradul_soft": "F5F1FC",
    "contrast_pink": "E83C96",
    "white": "FFFFFF",
    "grid": "DFDAF8",
}

CLASS_OPTIONS = ["Indefinido", "Procede", "No procede/Cambio Turno"]

def export_excel_cabify(df_main: pd.DataFrame, df_resumen: pd.DataFrame) -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = "Incidencias_por_comprobar"

    # Styles
    header_fill = PatternFill("solid", fgColor=CABIFY["moradul_dark"])
    header_font = Font(color=CABIFY["white"], bold=True)
    soft_fill = PatternFill("solid", fgColor=CABIFY["moradul_soft"])
    align = Alignment(vertical="center", wrap_text=True)
    thin = Side(style="thin", color=CABIFY["grid"])
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Write main sheet
    for r_idx, row in enumerate(dataframe_to_rows(df_main, index=False, header=True), start=1):
        ws.append(row)
        if r_idx == 1:
            for c_idx in range(1, len(row) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border
        else:
            # zebra
            if r_idx % 2 == 0:
                for c_idx in range(1, len(row) + 1):
                    ws.cell(row=r_idx, column=c_idx).fill = soft_fill
            for c_idx in range(1, len(row) + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.alignment = align
                cell.border = border

    # Freeze header + filter
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Column widths (aprox)
    widths = {
        "A": 12,  # Fecha
        "B": 18,  # Nombre
        "C": 18,  # Primer Apellido
        "D": 18,  # Segundo Apellido
        "E": 14,  # RUT
        "F": 14,  # Turno
        "G": 28,  # Especialidad
        "H": 28,  # Supervisor
        "I": 16,  # Tipo_Incidencia
        "J": 40,  # Detalle
        "K": 24,  # Clasificación Manual
    }
    for col_letter, w in widths.items():
        if col_letter <= chr(ord("A") + ws.max_column - 1):
            ws.column_dimensions[col_letter].width = w

    # Date format (Fecha in col A)
    # apply from row 2 to end
    for r in range(2, ws.max_row + 1):
        ws[f"A{r}"].number_format = "DD-MM-YYYY"

    # Dropdown validation for "Clasificación Manual"
    # Find column index by name
    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if "Clasificación Manual" in header:
        col_idx = header.index("Clasificación Manual") + 1
        col_letter = chr(ord("A") + col_idx - 1)
        dv = DataValidation(
            type="list",
            formula1=f'"{",".join(CLASS_OPTIONS)}"',
            allow_blank=True,
            showDropDown=True,
        )
        ws.add_data_validation(dv)
        dv.add(f"{col_letter}2:{col_letter}{ws.max_row}")

    # Second sheet: resumen
    ws2 = wb.create_sheet("Resumen_procede")
    for r_idx, row in enumerate(dataframe_to_rows(df_resumen, index=False, header=True), start=1):
        ws2.append(row)
        if r_idx == 1:
            for c_idx in range(1, len(row) + 1):
                cell = ws2.cell(row=r_idx, column=c_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = border
        else:
            for c_idx in range(1, len(row) + 1):
                cell = ws2.cell(row=r_idx, column=c_idx)
                cell.alignment = align
                cell.border = border

    ws2.freeze_panes = "A2"
    ws2.auto_filter.ref = ws2.dimensions
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 12

    # Save
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio


# ========================
# UI
# ========================
st.title("App Incidencias / Ausentismo / Asistencia")

with st.sidebar:
    st.header("Cargar archivos (Excel)")
    f_turnos = st.file_uploader("1) Codificación Turnos BUK", type=["xlsx"])
    f_activos = st.file_uploader("2) Trabajadores Activos + Turnos", type=["xlsx"])
    f_detalle = st.file_uploader("3) Detalle Turnos Colaboradores (Hoja1=Inasistencias, Hoja2=Asistencias)", type=["xlsx"])

    st.divider()
    st.subheader("Reglas")
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    st.caption("Déjalo vacío para no filtrar por área.")

if not all([f_turnos, f_activos, f_detalle]):
    st.info("Sube los 3 archivos para comenzar.")
    st.stop()

# ========================
# Load
# ========================
df_turnos = excel_to_df(f_turnos, 0)         # (por ahora no lo usamos en este MVP)
df_activos = excel_to_df(f_activos, 0)       # (por ahora no lo usamos en este MVP)
df_inasist = excel_to_df(f_detalle, 0)       # Hoja1
df_asist = excel_to_df(f_detalle, 1)         # Hoja2

# Normalize RUT
rut_col_inas = safe_col(df_inasist, "RUT")
rut_col_asist = safe_col(df_asist, "RUT")

if rut_col_inas:
    df_inasist["RUT_norm"] = df_inasist[rut_col_inas].apply(normalize_rut)
else:
    df_inasist["RUT_norm"] = ""

if rut_col_asist:
    df_asist["RUT_norm"] = df_asist[rut_col_asist].apply(normalize_rut)
else:
    df_asist["RUT_norm"] = ""

# Area filter (optional)
def maybe_filter_area(df, colname="Área"):
    if not only_area:
        return df
    c = safe_col(df, colname)
    if not c:
        return df
    return df[df[c].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()

df_inasist = maybe_filter_area(df_inasist, "Área")
df_asist = maybe_filter_area(df_asist, "Área")

# ========================
# Build Incidencias
# ========================
inc_rows = []

# ---- Asistencias: SOLO si hay Retraso o Salida Anticipada
col_fecha_entrada = safe_col(df_asist, "Fecha Entrada") or safe_col(df_asist, "Día")
df_asist["Fecha_base"] = df_asist[col_fecha_entrada].apply(try_parse_date_any) if col_fecha_entrada else pd.NaT

df_asist["Retraso_h"] = get_num(df_asist, "Retraso (horas)")
df_asist["SalidaAnt_h"] = get_num(df_asist, "Salida Anticipada (horas)")

mask_asist = (df_asist["Retraso_h"] > 0) | (df_asist["SalidaAnt_h"] > 0)
df_asist_inc = df_asist[mask_asist].copy()

df_asist_inc["Tipo_Incidencia"] = "Marcaje"
df_asist_inc["Detalle"] = (
    "Retraso_h=" + df_asist_inc["Retraso_h"].round(2).astype(str) +
    " | SalidaAnt_h=" + df_asist_inc["SalidaAnt_h"].round(2).astype(str)
)

# Map columns (existentes en tu archivo real)
cols_map_asist = {
    "Nombre": safe_col(df_asist_inc, "Nombre"),
    "Primer Apellido": safe_col(df_asist_inc, "Primer Apellido"),
    "Segundo Apellido": safe_col(df_asist_inc, "Segundo Apellido"),
    "RUT": safe_col(df_asist_inc, "RUT"),
    "Turno": safe_col(df_asist_inc, "Turno"),
    "Especialidad": safe_col(df_asist_inc, "Especialidad"),
    "Supervisor": safe_col(df_asist_inc, "Supervisor"),
}

df_asist_out = pd.DataFrame({
    "Fecha": df_asist_inc["Fecha_base"],
    "Nombre": df_asist_inc[cols_map_asist["Nombre"]] if cols_map_asist["Nombre"] else "",
    "Primer Apellido": df_asist_inc[cols_map_asist["Primer Apellido"]] if cols_map_asist["Primer Apellido"] else "",
    "Segundo Apellido": df_asist_inc[cols_map_asist["Segundo Apellido"]] if cols_map_asist["Segundo Apellido"] else "",
    "RUT": df_asist_inc[cols_map_asist["RUT"]] if cols_map_asist["RUT"] else df_asist_inc["RUT_norm"],
    "Turno": df_asist_inc[cols_map_asist["Turno"]] if cols_map_asist["Turno"] else "",
    "Especialidad": df_asist_inc[cols_map_asist["Especialidad"]] if cols_map_asist["Especialidad"] else "",
    "Supervisor": df_asist_inc[cols_map_asist["Supervisor"]] if cols_map_asist["Supervisor"] else "",
    "Tipo_Incidencia": df_asist_inc["Tipo_Incidencia"],
    "Detalle": df_asist_inc["Detalle"],
})

inc_rows.append(df_asist_out)

# ---- Inasistencias: las traemos como incidencias (detalle = motivo)
col_dia = safe_col(df_inasist, "Día")
df_inasist["Fecha_base"] = df_inasist[col_dia].apply(try_parse_date_any) if col_dia else pd.NaT

motivo_col = safe_col(df_inasist, "Motivo")
df_inasist["Detalle"] = df_inasist[motivo_col].astype(str) if motivo_col else ""

# Si quieres filtrar SOLO las que vienen como "-" (inasistencia preliminar), descomenta:
# if motivo_col:
#     df_inasist = df_inasist[df_inasist[motivo_col].astype(str).isin(["-", "Inasistencia", "nan", ""])]


cols_map_inas = {
    "Nombre": safe_col(df_inasist, "Nombre"),
    "Primer Apellido": safe_col(df_inasist, "Primer Apellido"),
    "Segundo Apellido": safe_col(df_inasist, "Segundo Apellido"),
    "RUT": safe_col(df_inasist, "RUT"),
    "Turno": safe_col(df_inasist, "Turno"),
    "Especialidad": safe_col(df_inasist, "Especialidad"),
    "Supervisor": safe_col(df_inasist, "Supervisor"),
}

df_inas_out = pd.DataFrame({
    "Fecha": df_inasist["Fecha_base"],
    "Nombre": df_inasist[cols_map_inas["Nombre"]] if cols_map_inas["Nombre"] else "",
    "Primer Apellido": df_inasist[cols_map_inas["Primer Apellido"]] if cols_map_inas["Primer Apellido"] else "",
    "Segundo Apellido": df_inasist[cols_map_inas["Segundo Apellido"]] if cols_map_inas["Segundo Apellido"] else "",
    "RUT": df_inasist[cols_map_inas["RUT"]] if cols_map_inas["RUT"] else df_inasist["RUT_norm"],
    "Turno": df_inasist[cols_map_inas["Turno"]] if cols_map_inas["Turno"] else "",
    "Especialidad": df_inasist[cols_map_inas["Especialidad"]] if cols_map_inas["Especialidad"] else "",
    "Supervisor": df_inasist[cols_map_inas["Supervisor"]] if cols_map_inas["Supervisor"] else "",
    "Tipo_Incidencia": "Inasistencia",
    "Detalle": df_inasist["Detalle"],
})

inc_rows.append(df_inas_out)

# Consolidate
df_incidencias = pd.concat(inc_rows, ignore_index=True)

# Clasificación manual default
df_incidencias["Clasificación Manual"] = "Indefinido"

# Date formatting (keep datetime for Excel, show nice in UI)
df_incidencias["Fecha"] = pd.to_datetime(df_incidencias["Fecha"], errors="coerce")

# Order + clean
desired_cols = [
    "Fecha", "Nombre", "Primer Apellido", "Segundo Apellido", "RUT",
    "Turno", "Especialidad", "Supervisor", "Tipo_Incidencia", "Detalle",
    "Clasificación Manual"
]
df_incidencias = df_incidencias[desired_cols].sort_values(["Fecha", "RUT"], na_position="last").reset_index(drop=True)

# ========================
# UI
# ========================
st.subheader("Reporte Total de Incidencias por Comprobar")

st.caption("Clasificación Manual: Indefinido / Procede / No procede/Cambio Turno")

edited = st.data_editor(
    df_incidencias,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Fecha": st.column_config.DateColumn(format="DD-MM-YYYY"),
        "Clasificación Manual": st.column_config.SelectboxColumn(
            options=CLASS_OPTIONS
        )
    }
)

st.subheader("Resumen (solo 'Procede')")
df_ok = edited[edited["Clasificación Manual"] == "Procede"].copy()

if len(df_ok) == 0:
    st.warning("Aún no hay incidencias marcadas como 'Procede'.")
    resumen = pd.DataFrame(columns=["Tipo_Incidencia", "Cantidad"])
else:
    resumen = (
        df_ok.groupby(["Tipo_Incidencia"], dropna=False)
            .size().reset_index(name="Cantidad")
            .sort_values("Cantidad", ascending=False)
    )
    st.dataframe(resumen, use_container_width=True)

# ========================
# Export
# ========================
st.subheader("Descarga")

excel_bytes = export_excel_cabify(edited, resumen)

st.download_button(
    "Descargar Excel consolidado (Cabify)",
    data=excel_bytes,
    file_name="reporte_incidencias_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
