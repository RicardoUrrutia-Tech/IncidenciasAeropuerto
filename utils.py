import pandas as pd
import re
from datetime import datetime, timedelta

def read_excel(uploaded_file) -> pd.DataFrame:
    return pd.read_excel(uploaded_file)

def _parse_time(s: str):
    # acepta "7:55:00" o "07:55" etc
    if pd.isna(s):
        return None
    s = str(s).strip()
    if not s:
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    return None

def _parse_shift_range(shift_str: str):
    """
    Devuelve (start_time, end_time, crosses_midnight)
    acepta "09:00-20:00", "7:00:00 - 15:00:00", etc.
    """
    if pd.isna(shift_str):
        return (None, None, False)
    s = str(shift_str).strip()
    if not s or s == "-" or s.lower() == "libre":
        return (None, None, False)

    # normaliza separadores
    s = s.replace("–", "-").replace("—", "-")
    s = re.sub(r"\s+", " ", s)
    s = s.replace(" - ", "-").replace(" -", "-").replace("- ", "-")

    if "-" not in s:
        return (None, None, False)

    a, b = [x.strip() for x in s.split("-", 1)]
    t1 = _parse_time(a)
    t2 = _parse_time(b)
    if not (t1 and t2):
        return (None, None, False)

    crosses = (datetime.combine(datetime.today(), t2) <= datetime.combine(datetime.today(), t1))
    return (t1, t2, crosses)

def build_shift_catalog(df_cod: pd.DataFrame) -> pd.DataFrame:
    # Espera columnas: Sigla, Horario, Tipo, Jornada
    df = df_cod.copy()
    # Limpieza básica de nombres de columnas
    df.columns = [str(c).strip() for c in df.columns]

    required = {"Sigla", "Horario"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Base codificación: faltan columnas {missing}")

    # parsea horas
    starts, ends, crosses = [], [], []
    for h in df["Horario"]:
        t1, t2, cr = _parse_shift_range(h)
        starts.append(t1)
        ends.append(t2)
        crosses.append(cr)

    df["HoraInicio"] = starts
    df["HoraFin"] = ends
    df["CruzaMedianoche"] = crosses

    # key normalizada
    df["Sigla_norm"] = df["Sigla"].astype(str).str.strip().str.upper()
    return df

def normalize_shift_to_range(value, shift_catalog: pd.DataFrame):
    """
    value puede ser Sigla o Horario.
    """
    if pd.isna(value):
        return (None, None, False, None)

    s = str(value).strip()
    if not s or s == "-" or s.lower() == "libre":
        return (None, None, False, None)

    # ¿Es sigla?
    key = s.upper()
    hit = shift_catalog[shift_catalog["Sigla_norm"] == key]
    if len(hit) == 1:
        row = hit.iloc[0]
        return (row["HoraInicio"], row["HoraFin"], bool(row["CruzaMedianoche"]), str(row["Sigla"]))

    # si no, intenta parsear como rango horario
    t1, t2, cr = _parse_shift_range(s)
    if t1 and t2:
        return (t1, t2, cr, None)

    return (None, None, False, None)

def prepare_activos_turnos(df_act: pd.DataFrame, shift_catalog: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte base ancha a larga:
    RUT + metadata + Fecha + TurnoOriginal + HoraInicioExp + HoraFinExp + CruzaMedianoche
    Aplica regla: solo desde la primera fecha con turno no vacío por trabajador.
    """
    df = df_act.copy()
    df.columns = [str(c).strip() for c in df.columns]

    fixed_cols = ["Nombre del Colaborador", "RUT", "Área", "Supervisor"]
    # tolerante: si cambian acentos o mayúsculas, igual intentamos
    # (si no están exactas, el usuario nos dirá y las mapeamos)
    meta = [c for c in df.columns if c in fixed_cols]

    date_cols = [c for c in df.columns if re.fullmatch(r"\d{2}-\d{2}-\d{4}", str(c).strip())]
    if not date_cols:
        raise ValueError("No encontré columnas fecha DD-MM-AAAA en 'Activos + Turnos'.")

    long = df.melt(id_vars=meta, value_vars=date_cols, var_name="Fecha", value_name="TurnoOriginal")
    long["Fecha"] = pd.to_datetime(long["Fecha"], format="%d-%m-%Y", errors="coerce")

    # normaliza turnos
    out = long.apply(
        lambda r: normalize_shift_to_range(r["TurnoOriginal"], shift_catalog),
        axis=1,
        result_type="expand",
    )
    out.columns = ["HoraInicioExp", "HoraFinExp", "CruzaMedianoche", "SiglaDetectada"]
    long = pd.concat([long, out], axis=1)

    # primer día válido por trabajador (turno con HoraInicioExp no nula)
    rut_col = "RUT" if "RUT" in long.columns else meta[0]
    first_valid = (
        long[long["HoraInicioExp"].notna()]
        .groupby(rut_col)["Fecha"]
        .min()
        .rename("PrimeraFechaActiva")
        .reset_index()
    )
    long = long.merge(first_valid, on=rut_col, how="left")
    long = long[long["Fecha"] >= long["PrimeraFechaActiva"]].copy()

    return long

def prepare_asistencias(df_asi: pd.DataFrame, shift_catalog: pd.DataFrame) -> pd.DataFrame:
    df = df_asi.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # parse fechas
    for col in ["Fecha Entrada", "Fecha Salida"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # parse horas
    for col in ["Hora Entrada", "Hora Salida"]:
        if col in df.columns:
            df[col] = df[col].apply(_parse_time)

    # normaliza turno declarado (opcional)
    if "Turno" in df.columns:
        out = df.apply(
            lambda r: normalize_shift_to_range(r["Turno"], shift_catalog),
            axis=1,
            result_type="expand",
        )
        out.columns = ["HoraInicioTurno", "HoraFinTurno", "CruzaMedianocheTurno", "SiglaDetectadaTurno"]
        df = pd.concat([df, out], axis=1)

    return df

def detect_incidencias(act_long: pd.DataFrame, asist: pd.DataFrame, df_det: pd.DataFrame,
                      tolerance_min: int = 5, manual_df=None) -> pd.DataFrame:
    """
    Regla base:
    - Para cada (RUT, Fecha) con turno esperado válido:
      buscamos marcaje de entrada/salida.
      Detectamos:
        - Sin marcaje entrada
        - Sin marcaje salida
        - Entrada tardía (minutos > tolerancia)
        - Salida anticipada (minutos > tolerancia)
    Nota: aquí dejamos la lógica simple y robusta; luego afinamos con tus datos reales.
    """
    df = act_long.copy()

    # claves
    if "RUT" not in df.columns or "RUT" not in asist.columns:
        raise ValueError("No encontré columna 'RUT' en una de las bases. Dime el nombre exacto y lo mapeo.")

    df = df[df["HoraInicioExp"].notna()].copy()

    # arma datetime esperado
    df["EntradaEsperada"] = df.apply(lambda r: datetime.combine(r["Fecha"].date(), r["HoraInicioExp"]), axis=1)
    df["SalidaEsperada"] = df.apply(
        lambda r: datetime.combine((r["Fecha"] + timedelta(days=1)).date(), r["HoraFinExp"])
        if r["CruzaMedianoche"]
        else datetime.combine(r["Fecha"].date(), r["HoraFinExp"]),
        axis=1,
    )

    # prepara asistencias con datetime reales (si vienen separadas)
    a = asist.copy()
    a = a[a["RUT"].notna()].copy()

    def _combine(fecha_col, hora_col):
        if fecha_col not in a.columns or hora_col not in a.columns:
            return pd.NaT
        dt = []
        for f, t in zip(a[fecha_col], a[hora_col]):
            if pd.isna(f) or t is None:
                dt.append(pd.NaT)
            else:
                dt.append(datetime.combine(pd.to_datetime(f).date(), t))
        return pd.to_datetime(dt, errors="coerce")

    a["EntradaRealDT"] = _combine("Fecha Entrada", "Hora Entrada")
    a["SalidaRealDT"] = _combine("Fecha Salida", "Hora Salida")

    # join aproximado: por RUT y por fecha de entrada esperada (día)
    # (más adelante afinamos si hay múltiples marcajes por día)
    a["FechaBase"] = pd.to_datetime(a["Fecha Entrada"], errors="coerce").dt.date
    df["FechaBase"] = df["Fecha"].dt.date

    merged = df.merge(
        a[["RUT", "FechaBase", "EntradaRealDT", "SalidaRealDT"]],
        on=["RUT", "FechaBase"],
        how="left",
    )

    tol = timedelta(minutes=tolerance_min)

    # detecta incidencias
    incidencias = []
    for _, r in merged.iterrows():
        tipos = []
        if pd.isna(r["EntradaRealDT"]):
            tipos.append("Sin marcaje entrada")
        else:
            if r["EntradaRealDT"] > (r["EntradaEsperada"] + tol):
                tipos.append("Entrada tardía")

        if pd.isna(r["SalidaRealDT"]):
            tipos.append("Sin marcaje salida")
        else:
            if r["SalidaRealDT"] < (r["SalidaEsperada"] - tol):
                tipos.append("Salida anticipada")

        for t in tipos:
            incidencias.append({
                "RUT": r["RUT"],
                "Fecha": r["Fecha"],
                "TurnoOriginal": r["TurnoOriginal"],
                "EntradaEsperada": r["EntradaEsperada"],
                "SalidaEsperada": r["SalidaEsperada"],
                "EntradaRealDT": r["EntradaRealDT"],
                "SalidaRealDT": r["SalidaRealDT"],
                "Tipo Incidencia": t,
                "Comprobación Incidencia": "Indefinido",
            })

    inc = pd.DataFrame(incidencias)

    # agrega manual si viene
    if manual_df is not None and len(manual_df) > 0:
        # asumimos que manual trae columnas compatibles; si no, lo ajustamos contigo
        inc = pd.concat([inc, manual_df], ignore_index=True)

    # ordena
    if "Fecha" in inc.columns:
        inc = inc.sort_values(["Fecha", "RUT", "Tipo Incidencia"]).reset_index(drop=True)

    return inc

def build_outputs(act_long: pd.DataFrame, asist: pd.DataFrame, incidencias_edit: pd.DataFrame) -> dict:
    out = {}

    # 2) incidencias por tipo (comprobadas)
    proc = incidencias_edit[incidencias_edit["Comprobación Incidencia"] == "Procede"].copy()
    if len(proc) == 0:
        out["incidencias_por_tipo"] = pd.DataFrame(columns=["Tipo Incidencia", "Q"])
    else:
        out["incidencias_por_tipo"] = proc.groupby("Tipo Incidencia").size().reset_index(name="Q").sort_values("Q", ascending=False)

    # 3) marcaje: conteos simples
    rep_m = pd.DataFrame({
        "Métrica": ["Registros marcaje (PBI)", "Con entrada real", "Con salida real"],
        "Valor": [len(asist), asist["Hora Entrada"].notna().sum() if "Hora Entrada" in asist.columns else asist["EntradaRealDT"].notna().sum(),
                  asist["Hora Salida"].notna().sum() if "Hora Salida" in asist.columns else asist["SalidaRealDT"].notna().sum()]
    })
    out["reporte_marcaje"] = rep_m

    # base de turnos esperados
    base = act_long[act_long["HoraInicioExp"].notna()].copy()
    base["FechaBase"] = base["Fecha"].dt.date

    # marca si tiene incidencia "Procede" en ese día
    proc_day = proc.copy()
    proc_day["FechaBase"] = pd.to_datetime(proc_day["Fecha"]).dt.date
    proc_flag = proc_day.groupby(["RUT", "FechaBase"]).size().reset_index(name="Inc_Procede_Q")
    base = base.merge(proc_flag, on=["RUT", "FechaBase"], how="left")
    base["Inc_Procede_Q"] = base["Inc_Procede_Q"].fillna(0)

    # 4) cumplimiento por trabajador (simple): días sin incidencias procede / total turnos
    cum = base.groupby("RUT").agg(
        Turnos=("FechaBase", "count"),
        Turnos_sin_incidencias_procede=("Inc_Procede_Q", lambda s: (s == 0).sum()),
        Turnos_con_incidencias_procede=("Inc_Procede_Q", lambda s: (s > 0).sum()),
    ).reset_index()
    cum["Pct_Cumplimiento"] = (cum["Turnos_sin_incidencias_procede"] / cum["Turnos"]).round(4)
    out["cumplimiento_trabajador"] = cum.sort_values("Pct_Cumplimiento")

    # 5) ausentismo (proxy): sin entrada real + turno esperado
    # (afinamos después con tus reglas finales)
    a = asist.copy()
    a["FechaBase"] = pd.to_datetime(a["Fecha Entrada"], errors="coerce").dt.date
    joined = base.merge(a[["RUT", "FechaBase", "EntradaRealDT"]], on=["RUT", "FechaBase"], how="left")
    aus = joined.groupby("RUT").agg(
        Turnos=("FechaBase", "count"),
        Sin_Entrada=("EntradaRealDT", lambda s: s.isna().sum()),
    ).reset_index()
    aus["Pct_Ausentismo"] = (aus["Sin_Entrada"] / aus["Turnos"]).round(4)
    out["ausentismo_trabajador"] = aus.sort_values("Pct_Ausentismo", ascending=False)

    # 6) asistencia diaria: % con entrada real
    daily = joined.groupby("FechaBase").agg(
        Turnos=("RUT", "count"),
        Con_Entrada=("EntradaRealDT", lambda s: s.notna().sum())
    ).reset_index()
    daily["Pct_Asistencia"] = (daily["Con_Entrada"] / daily["Turnos"]).round(4)
    out["asistencia_diaria"] = daily.sort_values("FechaBase")

    return out

def to_excel_bytes(obj, sheet_name="Sheet1", multi_sheet=False, filename_hint="out.xlsx") -> bytes:
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        if multi_sheet:
            # obj es dict[str, df]
            for k, v in obj.items():
                if isinstance(v, pd.DataFrame):
                    v.to_excel(writer, sheet_name=str(k)[:31], index=False)
        else:
            obj.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return buffer.getvalue()
