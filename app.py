import streamlit as st
import pandas as pd
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.datavalidation import DataValidation


st.set_page_config(page_title="Incidencias — Aeropuerto", layout="wide")
st.title("📌 Incidencias (Retraso / Salida anticipada) — Aeropuerto")


# ------------------------
# Helpers
# ------------------------
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

def get_num(df, col):
    if col not in df.columns:
        return pd.Series([0] * len(df))
    return pd.to_numeric(df[col], errors="coerce").fillna(0)

def read_excel_sheet(file, idx):
    return pd.read_excel(file, sheet_name=idx, engine="openpyxl")

def excel_col_letter(n: int) -> str:
    """1-indexed -> Excel column letters"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def to_cabify_excel_with_validation(df: pd.DataFrame, classification_col: str) -> BytesIO:
    # Paleta Cabify (según lo que pegaste)
    CABIFY_HEADER = "5B34AC"   # morado
    CABIFY_BORDER = "362065"
    CABIFY_ALTROW = "FAF8FE"   # casi blanco con tinte
    CABIFY_TEXT_WHITE = "FFFFFF"

    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Incidencias"

    # Estilos
    header_fill = PatternFill("solid", fgColor=CABIFY_HEADER)
    header_font = Font(color=CABIFY_TEXT_WHITE, bold=True)
    border_side = Side(style="thin", color=CABIFY_BORDER)
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    alt_fill = PatternFill("solid", fgColor=CABIFY_ALTROW)

    # Escribir DF (header + data)
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    max_row = ws.max_row
    max_col = ws.max_column

    # Header styling
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Auto filter
    ws.auto_filter.ref = f"A1:{excel_col_letter(max_col)}1"

    # Ajustes de ancho básicos + bordes + alternado
    for col_idx in range(1, max_col + 1):
        col_letter = excel_col_letter(col_idx)
        ws.column_dimensions[col_letter].width = 18

    # Fecha: formato corto (si existe)
    if "Fecha" in df.columns:
        fecha_col_idx = df.columns.get_loc("Fecha") + 1
    else:
        fecha_col_idx = None

    for row in ws.iter_rows(min_row=2, max_row=max_row, max_col=max_col):
        excel_row = row[0].row

        # Alternado
        if excel_row % 2 == 0:
            for cell in row:
                cell.fill = alt_fill

        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

        # Formato fecha corto
        if fecha_col_idx is not None:
            c = row[fecha_col_idx - 1]
            c.number_format = "dd-mm-yyyy"

    # ------------------------------
    # Data validation: Clasificación Manual (lista)
    # ------------------------------
    if classification_col in df.columns:
        options = ["Seleccionar", "Injustificada", "Permiso", "No Procede - Cambio de Turno"]
        formula = '"' + ",".join(options) + '"'

        dv = DataValidation(type="list", formula1=formula, allow_blank=False)
        col_idx = df.columns.get_loc(classification_col) + 1
        col_letter = excel_col_letter(col_idx)

        # aplicar desde fila 2 hasta max fila (y de paso dejemos muchas filas por si agregan)
        dv.add(f"{col_letter}2:{col_letter}1048576")
        ws.add_data_validation(dv)

    wb.save(output)
    output.seek(0)
    return output


# ------------------------
# UI: upload
# ------------------------
with st.sidebar:
    st.header("📤 Cargar archivo")
    st.caption("Sube el Excel **Detalle Turnos Colaboradores** (Hoja1=Inasistencias, Hoja2=Asistencias).")
    f_detalle = st.file_uploader("Detalle Turnos Colaboradores (Excel)", type=["xlsx"])

    st.divider()
    st.subheader("Regla")
    st.caption("Solo se consideran incidencias cuando exista Retraso o Salida Anticipada (> 0).")

if not f_detalle:
    st.info("Sube el Excel **Detalle Turnos Colaboradores** para comenzar.")
    st.stop()


# ------------------------
# Load sheets
# ------------------------
df_inasist = read_excel_sheet(f_detalle, 0)  # Hoja 1 (no usada en el filtro final)
df_asist = read_excel_sheet(f_detalle, 1)    # Hoja 2 (la que usaremos)

# Normalizaciones mínimas
if "RUT" in df_asist.columns:
    df_asist["RUT"] = df_asist["RUT"].astype(str)
    df_asist["RUT_norm"] = df_asist["RUT"].apply(normalize_rut)
else:
    st.error("No encuentro la columna 'RUT' en la Hoja 2 (Asistencias).")
    st.stop()

# Fecha base = Fecha Entrada (según tu definición)
if "Fecha Entrada" in df_asist.columns:
    df_asist["Fecha"] = df_asist["Fecha Entrada"].apply(try_parse_date_any).dt.date
elif "Día" in df_asist.columns:
    df_asist["Fecha"] = df_asist["Día"].apply(try_parse_date_any).dt.date
else:
    st.error("No encuentro 'Fecha Entrada' ni 'Día' en la Hoja 2 (Asistencias).")
    st.stop()

# Numéricos
df_asist["Retraso_h"] = get_num(df_asist, "Retraso (horas)")
df_asist["SalidaAnt_h"] = get_num(df_asist, "Salida Anticipada (horas)")

# ------------------------
# FILTRO PRINCIPAL: solo incidencias reales
# ------------------------
mask = (df_asist["Retraso_h"] > 0) | (df_asist["SalidaAnt_h"] > 0)
df_inc = df_asist[mask].copy()

if df_inc.empty:
    st.warning("No hay incidencias: ningún registro con Retraso>0 o Salida Anticipada>0.")
    st.stop()

# Tipo_Incidencia + Detalle
def tipo_incidencia(row):
    r = float(row.get("Retraso_h", 0) or 0)
    s = float(row.get("SalidaAnt_h", 0) or 0)
    if r > 0 and s > 0:
        return "Retraso + Salida anticipada"
    if r > 0:
        return "Retraso"
    return "Salida anticipada"

df_inc["Tipo_Incidencia"] = df_inc.apply(tipo_incidencia, axis=1)
df_inc["Detalle"] = (
    "Retraso_h=" + df_inc["Retraso_h"].round(2).astype(str)
    + " | SalidaAnt_h=" + df_inc["SalidaAnt_h"].round(2).astype(str)
)

# Clasificación Manual default
CLASIF_COL = "Clasificación Manual"
if CLASIF_COL not in df_inc.columns:
    df_inc[CLASIF_COL] = "Seleccionar"
else:
    df_inc[CLASIF_COL] = df_inc[CLASIF_COL].fillna("Seleccionar").astype(str).replace("", "Seleccionar")

# ------------------------
# Salida con SOLO los campos solicitados
# ------------------------
wanted = [
    "Fecha",
    "Nombre",
    "Primer Apellido",
    "Segundo Apellido",
    "RUT",
    "Turno",
    "Especialidad",
    "Supervisor",
    "Tipo_Incidencia",
    "Detalle",
    CLASIF_COL,
]
missing = [c for c in wanted if c not in df_inc.columns]
if missing:
    st.error(f"Faltan columnas en tu Hoja 2 (Asistencias): {missing}")
    st.stop()

df_out = df_inc[wanted].copy()

# Orden por Fecha/RUT
df_out = df_out.sort_values(["Fecha", "RUT"], na_position="last")


# ------------------------
# UI table + editor
# ------------------------
st.subheader("🧾 Incidencias detectadas (solo si hay Retraso o Salida Anticipada)")
st.caption("La columna **Clasificación Manual** se edita aquí y se exporta con selector en Excel.")

edited = st.data_editor(
    df_out,
    use_container_width=True,
    num_rows="fixed",
    column_config={
        CLASIF_COL: st.column_config.SelectboxColumn(
            options=["Seleccionar", "Injustificada", "Permiso", "No Procede - Cambio de Turno"]
        )
    }
)

# ------------------------
# Export to Excel (Cabify style + dropdown)
# ------------------------
st.subheader("⬇️ Descargar Excel")
excel_bytes = to_cabify_excel_with_validation(edited, classification_col=CLASIF_COL)

st.download_button(
    "Descargar Excel (Cabify + selector)",
    data=excel_bytes,
    file_name="incidencias_aeropuerto.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

