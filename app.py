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

def maybe_filter_area(df, only_area: str, col="Área"):
    if only_area and col in df.columns:
        return df[df[col].astype(str).str.upper().str.contains(only_area.upper(), na=False)].copy()
    return df

def safe_col(df: pd.DataFrame, name: str) -> str | None:
    """Busca columna por nombre exacto; si no existe, None."""
    return name if name in df.columns else None

def build_export_excel_openpyxl(df_main: pd.DataFrame, df_resumen: pd.DataFrame) -> BytesIO:
    """
    Genera Excel con estilo + data validation (dropdown) en 'Clasificación Manual'
    usando openpyxl (sin xlsxwriter).
    """
    output = BytesIO()
    wb = Workbook()

    # Paleta Cabify (según lo que indicaste)
    cabify_dark = "1F123F"     # header
    cabify_light = "F5F1FC"    # alternado
    cabify_mid = "DFDAF8"
    cabify_accent = "E83C96"   # contraste

    # Estilos
    header_fill = PatternFill("solid", fgColor=cabify_dark)
    header_font = Font(color="FFFFFF", bold=True)
    alt_fill = PatternFill("solid", fgColor=cabify_light)
    base_fill = PatternFill("solid", fgColor="FFFFFF")

    border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    def write_df(ws, df: pd.DataFrame, sheet_title: str, date_cols: list[str] = None, add_dropdown: bool = False):
        ws.title = sheet_title[:31]

        # Escribir dataframe
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
            ws.append(row)

        # Header
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        # Body styling + bordes + alternado
        max_row = ws.max_row
        max_col = ws.max_column

        for r in range(2, max_row + 1):
            is_alt = (r % 2 == 0)
            row_fill = alt_fill if is_alt else base_fill
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.fill = row_fill
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=True)

        # Freeze panes + autofilter
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Formato fecha corta
        if date_cols:
            for dc in date_cols:
                if dc in df.columns:
                    col_idx = df.columns.get_loc(dc) + 1
                    for r in range(2, max_row + 1):
                        cell = ws.cell(row=r, column=col_idx)
                        cell.number_format = "DD-MM-YYYY"

        # Anchos aproximados
        for col_idx in range(1, max_col + 1):
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            header_val = ws.cell(row=1, column=col_idx).value or ""
            # ancho basado en header (con mínimo)
            width = max(12, min(45, len(str(header_val)) + 8))
            ws.column_dimensions[col_letter].width = width

        # Dropdown en Clasificación Manual
        if add_dropdown and "Clasificación Manual" in df.columns:
            options = "Seleccionar,Injustificada,Permiso,No Procede - Cambio de Turno"
            dv = DataValidation(type="list", formula1=f'"{options}"', allow_blank=False)
            col_idx = df.columns.get_loc("Clasificación Manual") + 1
            col_letter = ws.cell(row=1, column=col_idx).column_letter

            # Aplicar a un rango amplio (hasta max_row)
            dv.add(f"{col_letter}2:{col_letter}{max_row}")
            ws.add_data_validation(dv)

    # Hoja 1: Incidencias
    ws1 = wb.active
    write_df(
        ws1,
        df_main,
        "Incidencias_por_comprobar",
        date_cols=["Fecha"],
        add_dropdown=True
    )

    # Hoja 2: Resumen
    ws2 = wb.create_sheet("Resumen")
    write_df(ws2, df_resumen, "Resumen", date_cols=[])

    wb.save(output)
    output.seek(0)
    return output


# =========================
# UI - Uploads
# =========================
with st.sidebar:
    st.header("Cargar archivos (Excel)")

    f_turnos = st.file_uploader("1) Codificación Turnos BUK", type=["xlsx"])
    f_activos = st.file_uploader("2) Trabajadores Activos + Turnos", type=["xlsx"])
    f_detalle = st.file_uploader("3) Detalle Turnos Colaboradores (Hoja1 Inasistencias + Hoja2 Asistencias)", type=["xlsx"])

    st.divider()
    st.subheader("Opcionales")
    only_area = st.text_input("Filtrar Área (opcional)", value="AEROPUERTO")
    st.caption("Déjalo vacío para no filtrar por Área.")


if not all([f_turnos, f_activos, f_detalle]):
    st.info("Sube los 3 archivos para comenzar.")
    st.stop()


# =========================
# Load
# =========================
df_turnos = excel_to_df(f_turnos, 0)       # (MVP: no se usa todavía, pero se carga)
df_activos = excel_to_df(f_activos, 0)     # (MVP: se carga; lo usaremos luego para cruces si lo pides)
df_inasist = excel_to_df(f_detalle, 0)     # Hoja1
df_asist = excel_to_df(f_detalle, 1)       # Hoja2

# Normalizaciones básicas
for df in [df_inasist, df_asist]:
    if "RUT" in df.columns:
        df["RUT_norm"] = df["RUT"].apply(normalize_rut)
    else:
        df["RUT_norm"] = ""

df_inasist = maybe_filter_area(df_inasist, only_area, "Área")
df_asist = maybe_filter_area(df_asist, only_area, "Área")


# =========================
# Construcción de incidencias
# =========================
inc_rows = []

# --------
# Asistencias -> Incidencias SOLO si hay Retraso o Salida Anticipada
# --------
def get_num(df, col):
    if col not in df.columns:
        return pd.Series([0] * len(df))
    return pd.to_numeric(df[col], errors="coerce").fillna(0)

df_asist["Retraso_h"] = get_num(df_asist, "Retraso (horas)")
df_asist["SalidaAnt_h"] = get_num(df_asist, "Salida Anticipada (horas)")

# Fecha base en asistencias: Fecha Entrada
if "Fecha Entrada" in df_asist.columns:
    df_asist["Fecha_base"] = df_asist["Fecha Entrada"].apply(try_parse_date_any)
else:
    df_asist["Fecha_base"] = pd.NaT

mask_asist = (df_asist["Retraso_h"] > 0) | (df_asist["SalidaAnt_h"] > 0)
df_asist_inc = df_asist[mask_asist].copy()

df_asist_inc["Tipo_Incidencia"] = "Marcaje/Turno"
df_asist_inc["Detalle"] = (
    "Retraso_h=" + df_asist_inc["Retraso_h"].astype(str) +
    " | SalidaAnt_h=" + df_asist_inc["SalidaAnt_h"].astype(str)
)

# Selección de campos requeridos
cols_needed_asist = {
    "Fecha": "Fecha_base",
    "Nombre": "Nombre",
    "Primer Apellido": "Primer Apellido",
    "Segundo Apellido": "Segundo Apellido",
    "RUT": "RUT",
    "Turno": "Turno",
    "Especialidad": "Especialidad",
    "Supervisor": "Supervisor",
    "Tipo_Incidencia": "Tipo_Incidencia",
    "Detalle": "Detalle",
}

tmp = pd.DataFrame()
for out_col, in_col in cols_needed_asist.items():
    if in_col in df_asist_inc.columns:
        tmp[out_col] = df_asist_inc[in_col]
    else:
        tmp[out_col] = ""

inc_rows.append(tmp)

# --------
# Inasistencias (Hoja1) -> incluir las inasistencias preliminares (Motivo "-" o vacío o "INASISTENCIA")
# --------
if "Día" in df_inasist.columns:
    df_inasist["Fecha_base"] = df_inasist["Día"].apply(try_parse_date_any)
else:
    df_inasist["Fecha_base"] = pd.NaT

motivo = df_inasist["Motivo"].astype(str).str.strip().str.upper() if "Motivo" in df_inasist.columns else pd.Series([""]*len(df_inasist))
mask_inasist = motivo.isin(["-", "", "INASISTENCIA"])

df_inasist_inc = df_inasist[mask_inasist].copy()
df_inasist_inc["Tipo_Incidencia"] = "Inasistencia"

obs_perm = df_inasist_inc["Observación Permiso"].astype(str) if "Observación Permiso" in df_inasist_inc.columns else ""
mot = df_inasist_inc["Motivo"].astype(str) if "Motivo" in df_inasist_inc.columns else ""
df_inasist_inc["Detalle"] = ("Motivo=" + mot + " | Obs=" + obs_perm).astype(str)

cols_needed_inasist = {
    "Fecha": "Fecha_base",
    "Nombre": "Nombre",
    "Primer Apellido": "Primer Apellido",
    "Segundo Apellido": "Segundo Apellido",
    "RUT": "RUT",
    "Turno": "Turno",
    "Especialidad": "Especialidad",
    "Supervisor": "Supervisor",
    "Tipo_Incidencia": "Tipo_Incidencia",
    "Detalle": "Detalle",
}

tmp2 = pd.DataFrame()
for out_col, in_col in cols_needed_inasist.items():
    if in_col in df_inasist_inc.columns:
        tmp2[out_col] = df_inasist_inc[in_col]
    else:
        tmp2[out_col] = ""

inc_rows.append(tmp2)

# Consolidado
df_incidencias = pd.concat(inc_rows, ignore_index=True)

# Fecha como date (sin hora)
df_incidencias["Fecha"] = pd.to_datetime(df_incidencias["Fecha"], errors="coerce").dt.date

# Clasificación Manual (default)
if "Clasificación Manual" not in df_incidencias.columns:
    df_incidencias["Clasificación Manual"] = "Seleccionar"

# Orden final exacto (solo campos pedidos)
final_cols = [
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
    "Clasificación Manual",
]
df_incidencias = df_incidencias[final_cols].sort_values(["Fecha", "RUT"], na_position="last")


# =========================
# UI: tabla editable
# =========================
st.subheader("Reporte Total de Incidencias por Comprobar")

edited = st.data_editor(
    df_incidencias,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Clasificación Manual": st.column_config.SelectboxColumn(
            options=[
                "Seleccionar",
                "Injustificada",
                "Permiso",
                "No Procede - Cambio de Turno",
            ]
        )
    }
)

# =========================
# Resumen (opcional)
# =========================
st.subheader("Resumen (solo 'Injustificada')")
df_ok = edited[edited["Clasificación Manual"] == "Injustificada"].copy()

if len(df_ok) == 0:
    st.warning("Aún no hay registros marcados como 'Injustificada'.")
    resumen = pd.DataFrame(columns=["Tipo_Incidencia", "Cantidad"])
else:
    resumen = (
        df_ok.groupby(["Tipo_Incidencia"], dropna=False)
            .size().reset_index(name="Cantidad")
            .sort_values("Cantidad", ascending=False)
    )
    st.dataframe(resumen, use_container_width=True)

# =========================
# Export
# =========================
st.subheader("Descarga")
excel_bytes = build_export_excel_openpyxl(edited, resumen)

st.download_button(
    "Descargar Excel consolidado (con selector + estilo Cabify)",
    data=excel_bytes,
    file_name="reporte_incidencias_consolidado.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

