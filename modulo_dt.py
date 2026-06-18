"""
modulo_dt.py
Procesamiento del archivo de la Dirección del Trabajo (DT)
para generación del archivo de liquidaciones en detalle Rex+.
"""

import re
import io
import os
import calendar
import pandas as pd
import numpy as np
import streamlit as st
from io import StringIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime


# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
COL_RUT           = "Rut trabajador(1101)"
COL_DIAS_TRAB     = "Nro días trabajados en el mes(1115)"
COL_DIAS_LIC      = "Nro días de licencia médica en el mes(1116)"
COL_DIAS_VAC      = "Nro días de vacaciones en el mes(1117)"
COL_SUELDO        = "Sueldo(2101)"
COL_SALUD7        = "Cotización obligatoria salud 7%(3143)"
COL_SALUD_VOL     = "Cotización voluntaria para salud(3144)"
COL_AFP           = "Cotización obligatoria previsional (AFP o IPS)(3141)"
COL_CES_TRAB      = "Cotización AFC - trabajador(3151)"
COL_APVI_MOD_B    = "Cotización APVi Mod B hasta UF50(3156)"
COL_TRAB_PESADO   = "Cotización adicional trabajo pesado - trabajador(3154)"
COL_REBAJA_ZONA   = "Rebaja zona extrema DL 889 (3167)"
COL_AFP_INST      = "AFP(1141)"
COL_ISAPRE_INST   = "FONASA - ISAPRE(1143)"
COL_MUTUAL_INST   = "Org. administrador ley 16.744(1152)"
COL_CCAF_INST     = "CCAF(1110)"

CONCEPTOS_AFECTO_AFP = {"mutual", "sis", "trabajoPesaEmpl", "afp", "isapre"}
CONCEPTOS_AFECTO_CES = {"cesAporteCi", "cesAporteSol", "cesEmpleado"}
CONCEPTOS_ID_AFP     = {"sis", "afp", "trabajoPesaEmpl", "cesEmpleado", "cesAporteSol", "cesAporteCi"}


# ─────────────────────────────────────────────
# LECTURA DEL CSV DT (encabezado variable)
# ─────────────────────────────────────────────
def leer_csv_dt(file_obj):
    """Lee el CSV de la DT detectando automáticamente la posición del encabezado."""
    for enc in ("utf-8", "latin-1", "utf-8-sig", "cp1252"):
        try:
            file_obj.seek(0)
            raw = file_obj.read().decode(enc)
            break
        except Exception:
            continue
    else:
        raise ValueError("No se pudo decodificar el archivo CSV.")

    lines = raw.splitlines()
    header_idx = None
    for i, line in enumerate(lines):
        if "Rut trabajador" in line:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("No se encontró la línea de encabezado en el archivo.")

    data_lines = [lines[header_idx]] + [l for i, l in enumerate(lines) if i != header_idx and l.strip()]
    content = "\n".join(data_lines)
    df = pd.read_csv(StringIO(content), sep=";", dtype=str)
    for col in df.columns:
        converted = pd.to_numeric(df[col], errors="coerce")
        if converted.notna().sum() > df[col].notna().sum() * 0.5:
            df[col] = converted
    return df


# ─────────────────────────────────────────────
# EXTRACCIÓN DE FECHA DESDE NOMBRE DE ARCHIVO
# ─────────────────────────────────────────────
def extraer_fecha_dt(nombre_archivo):
    """
    Intenta extraer yyyy-mm del nombre del archivo.
    Patrones soportados: 'enero_2025', '2025_01', '202501', 'enero-2025', etc.
    Retorna (fecha_str, True) si se encontró, (None, False) si no.
    """
    meses_es = {
        "enero": "01", "febrero": "02", "marzo": "03", "abril": "04",
        "mayo": "05", "junio": "06", "julio": "07", "agosto": "08",
        "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12"
    }

    nombre = nombre_archivo.lower()

    # Patrón: mes_nombre + año  ej: enero_2025, enero-2025
    for mes_nom, mes_num in meses_es.items():
        patron = rf"{mes_nom}[_\-\s](\d{{4}})"
        m = re.search(patron, nombre)
        if m:
            return f"{m.group(1)}-{mes_num}", True
        # año + mes_nombre  ej: 2025_enero
        patron2 = rf"(\d{{4}})[_\-\s]{mes_nom}"
        m2 = re.search(patron2, nombre)
        if m2:
            return f"{m2.group(1)}-{mes_num}", True

    # Patrón: yyyymm  ej: 202501
    m = re.search(r"(20\d{2})(0[1-9]|1[0-2])", nombre)
    if m:
        return f"{m.group(1)}-{m.group(2)}", True

    # Patrón: yyyy-mm o yyyy_mm
    m = re.search(r"(20\d{2})[_\-](0[1-9]|1[0-2])", nombre)
    if m:
        return f"{m.group(1)}-{m.group(2)}", True

    return None, False


# ─────────────────────────────────────────────
# CARGA DE REFERENCIAS DT
# ─────────────────────────────────────────────
def cargar_referencias_dt(data_dir="data"):
    """Carga todos los archivos de referencia necesarios para el módulo DT."""
    refs = {}
    errores = []
    archivos = {
        "equiv_conceptos":    "equiv_conceptos.xlsx",
        "parametros":         "parametrosMesuales.xlsx",
        "inst_afp":           "inst_afp.xlsx",
        "inst_mutuales":      "inst_mutuales.xlsx",
        "inst_salud":         "inst_salud.xlsx",
        "inst_cajas":         "inst_cajas.xlsx",
        "listado_empresas":   "listado_empresas.xlsx",
    }
    for key, fname in archivos.items():
        path = os.path.join(data_dir, fname)
        if os.path.exists(path):
            refs[key] = pd.read_excel(path)
        else:
            errores.append(fname)
    return refs, errores


# ─────────────────────────────────────────────
# CARGA LISTADO EMPLEADOS (con encabezado en fila 1)
# ─────────────────────────────────────────────
def cargar_empleados(file_obj):
    """
    Lee listado_empleados.xlsx que tiene encabezado real en la fila 1 (índice 0 es título).
    Retorna DataFrame limpio con columnas reales.
    """
    df = pd.read_excel(file_obj, header=1)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def cargar_empresas(file_obj_or_df):
    """
    Lee listado_empresas.xlsx que tiene encabezado real en la fila 1.
    Acepta path, file object o DataFrame ya cargado.
    """
    if isinstance(file_obj_or_df, pd.DataFrame):
        return file_obj_or_df
    df = pd.read_excel(file_obj_or_df, header=1)
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ─────────────────────────────────────────────
# RESOLUCIÓN DE CONTRATOS
# ─────────────────────────────────────────────
def resolver_contrato(df_empleados, rut, fecha_proceso):
    """
    Dado un RUT y la fecha de proceso (yyyy-mm), determina qué contrato aplica.

    Retorna:
        (contrato, empresa_codigo, ok, motivo)
        - ok=True  → se encontró exactamente un contrato válido
        - ok=False → ninguno o más de uno coincide (motivo indica el problema)
    """
    if "Rut" not in df_empleados.columns:
        return "", "", False, "Sin columna Rut"

    emp_rows = df_empleados[df_empleados["Rut"] == rut]
    if emp_rows.empty:
        return "", "", False, "RUT no encontrado en listado de empleados"

    # Si tiene un solo contrato, retornar directamente
    if len(emp_rows) == 1:
        row = emp_rows.iloc[0]
        contrato = row.get("Contrato", "")
        empresa  = str(row.get("Empresa", "")).strip()
        return contrato, empresa, True, ""

    # Múltiples contratos → resolver por intervalo de fechas
    try:
        fp_date = datetime.strptime(fecha_proceso, "%Y-%m")
    except Exception:
        return "", "", False, "Fecha de proceso inválida"

    col_inicio  = next((c for c in ["Fecha Inicio contrato", "Fecha inicio contrato",
                                     "Fecha inicio", "FechaInicio"] if c in emp_rows.columns), None)
    col_termino = next((c for c in ["Fecha término contrato", "Fecha termino contrato",
                                     "Fecha término", "FechaTermino"] if c in emp_rows.columns), None)

    if col_inicio is None:
        return "", "", False, "No se encontró columna de fecha inicio contrato"

    candidatos = []
    for _, cr in emp_rows.iterrows():
        try:
            f_inicio = pd.to_datetime(cr[col_inicio])
            if pd.isna(f_inicio):
                continue

            # Fecha término: si está vacía = contrato indefinido
            f_termino = None
            if col_termino:
                ft = pd.to_datetime(cr[col_termino], errors="coerce")
                if not pd.isna(ft):
                    f_termino = ft

            # Verificar si la fecha de proceso cae en el intervalo
            if f_inicio <= fp_date:
                if f_termino is None or fp_date <= f_termino:
                    candidatos.append(cr)
        except Exception:
            continue

    if len(candidatos) == 1:
        contrato = candidatos[0].get("Contrato", "")
        empresa  = str(candidatos[0].get("Empresa", "")).strip()
        return contrato, empresa, True, ""
    elif len(candidatos) == 0:
        return "", "", False, f"Ningún contrato aplica para el período {fecha_proceso}"
    else:
        return "", "", False, f"Más de un contrato aplica para el período {fecha_proceso} ({len(candidatos)} contratos)"


def construir_log_contratos(df_empleados, ruts_problema, fecha_proceso, motivos):
    """Construye el DataFrame de log para trabajadores con problema de contratos."""
    cols = [c for c in ["Rut", "Nombre", "Empresa", "Contrato",
                         "Fecha Inicio contrato", "Fecha término contrato"]
            if c in df_empleados.columns]
    detalle = df_empleados[df_empleados["Rut"].isin(ruts_problema)][cols].copy()
    detalle["Motivo"] = detalle["Rut"].map(motivos)
    detalle["Período"] = fecha_proceso
    return detalle


# ─────────────────────────────────────────────
# LOOKUP HELPERS
# ─────────────────────────────────────────────
def lookup(df, col_busca, valor, col_trae, default=""):
    """Busca valor en col_busca y retorna col_trae."""
    if df is None or df.empty:
        return default
    try:
        valor_num = pd.to_numeric(valor, errors="coerce")
        mask = df[col_busca] == (valor_num if not pd.isna(valor_num) else valor)
        if mask.any():
            return df.loc[mask.idxmax(), col_trae]
    except Exception:
        pass
    return default


def safe_num(val, default=0):
    """Convierte a número de forma segura."""
    try:
        v = float(val)
        return v if not np.isnan(v) else default
    except Exception:
        return default


# ─────────────────────────────────────────────
# GENERACIÓN DE FILAS DE SALIDA
# ─────────────────────────────────────────────
def generar_filas_dt(df, fecha_proceso, refs, df_empleados, df_empresas_externo=None):
    """
    Genera las filas del archivo de salida desde el CSV de la DT.
    df_empresas_externo: DataFrame de empresas ya cargado correctamente (prioritario).
    """
    filas = []

    equiv        = refs.get("equiv_conceptos", pd.DataFrame())
    params_df    = refs.get("parametros", pd.DataFrame())
    inst_afp     = refs.get("inst_afp", pd.DataFrame())
    inst_mutuales= refs.get("inst_mutuales", pd.DataFrame())
    inst_salud   = refs.get("inst_salud", pd.DataFrame())
    inst_cajas   = refs.get("inst_cajas", pd.DataFrame())

    # Usar df_empresas_externo si se pasó, si no intentar desde refs
    if df_empresas_externo is not None and not df_empresas_externo.empty:
        df_empresas = df_empresas_externo.copy()
    else:
        empresas_raw = refs.get("listado_empresas", pd.DataFrame())
        if isinstance(empresas_raw, pd.DataFrame) and "Empresa" in empresas_raw.columns:
            df_empresas = empresas_raw.copy()
        else:
            df_empresas = cargar_empresas(empresas_raw)

    # Normalizar columnas Nombre y Empresa
    if "Nombre" in df_empresas.columns:
        df_empresas["Nombre"] = df_empresas["Nombre"].astype(str).str.strip()
    if "Empresa" in df_empresas.columns:
        df_empresas["Empresa"] = df_empresas["Empresa"].astype(str).str.strip()

    # ── Parámetros del mes ──
    tope_salud = 0
    tope_afp   = 0
    tope_ces   = 0
    row_params = pd.DataFrame()
    if not params_df.empty and "mes_Proc" in params_df.columns:
        params_df["mes_Proc"] = params_df["mes_Proc"].astype(str).str.strip()
        row_params = params_df[params_df["mes_Proc"] == fecha_proceso]
        if not row_params.empty:
            tope_salud = safe_num(row_params.iloc[0].get("topeSalud_pesos", 0))
            tope_afp   = safe_num(row_params.iloc[0].get("topeImp_pesos_afp", 0))
            tope_ces   = safe_num(row_params.iloc[0].get("topeCes_pesos", 0))

    # ── Mapa de equivalencias: cod_lre → concepto_detalle + Tipo ──
    equiv_map  = {}  # col_csv → concepto_detalle
    tipo_map   = {}  # concepto_detalle → Tipo
    if not equiv.empty and "cod_lre" in equiv.columns and "concepto_detalle" in equiv.columns:
        for _, er in equiv.iterrows():
            col_csv   = str(er["cod_lre"]).strip()
            concepto  = str(er["concepto_detalle"]).strip()
            tipo      = str(er.get("Tipo", "")).strip()
            # Solo mapear la primera aparición para evitar duplicados de isapre
            if col_csv not in equiv_map:
                equiv_map[col_csv] = concepto
            tipo_map[concepto] = tipo

    # Columnas del CSV que tienen equivalencia (excluir COL_SALUD_VOL para isapre, se suma manualmente)
    cols_conceptos = [c for c in df.columns if c in equiv_map and c != COL_SALUD_VOL]

    # ── Registro de problemas de contrato (para log) ──
    ruts_problema = {}   # rut → motivo
    filas = []

    for _, row in df.iterrows():
        rut = str(row.get(COL_RUT, "")).strip()

        # ── Resolver contrato ──
        numero_contrato, empresa_codigo, ok, motivo = resolver_contrato(
            df_empleados, rut, fecha_proceso
        )
        if not ok:
            ruts_problema[rut] = motivo
            continue

        # ── Lookup empresa → busca por Nombre, trae código Empresa ──
        empresa_salida = empresa_codigo
        if empresa_codigo and not df_empresas.empty and "Nombre" in df_empresas.columns:
            empresa_codigo_strip = str(empresa_codigo).strip()
            df_empresas["Nombre"] = df_empresas["Nombre"].astype(str).str.strip()
            emp2 = df_empresas[df_empresas["Nombre"] == empresa_codigo_strip]
            if not emp2.empty:
                empresa_salida = str(emp2.iloc[0].get("Empresa", empresa_codigo)).strip()

        # ── Valores base del trabajador ──
        dias_trab   = safe_num(row.get(COL_DIAS_TRAB, 0))
        dias_lic    = safe_num(row.get(COL_DIAS_LIC, 0))
        dias_vac    = safe_num(row.get(COL_DIAS_VAC, 0))
        sueldo      = safe_num(row.get(COL_SUELDO, 0))
        rebaja_zona = safe_num(row.get(COL_REBAJA_ZONA, 0))

        monto_init = round((sueldo / dias_trab) * 30, 0) if dias_trab > 0 else 0

        # ── Suma haberes afectos (para Afecto) ──
        conceptos_haber_afecto = [c for c in cols_conceptos
                                   if tipo_map.get(equiv_map.get(c, ""), "") == "Haber afecto"]
        suma_haber_afecto = sum(safe_num(row.get(c, 0)) for c in conceptos_haber_afecto)

        # ── Suma haberes exentos (para Rentas no gravadas) ──
        conceptos_haber_exento = [c for c in cols_conceptos
                                   if tipo_map.get(equiv_map.get(c, ""), "") == "Haber exento"]
        suma_haber_exento = sum(safe_num(row.get(c, 0)) for c in conceptos_haber_exento)

        # ── Monto isapre (caso especial: suma dos columnas) ──
        monto_isapre = safe_num(row.get(COL_SALUD7, 0)) + safe_num(row.get(COL_SALUD_VOL, 0))

        # ── Total rebajas por LLSS (solo para concepto impuesto) ──
        salud_trab = safe_num(row.get(COL_SALUD7, 0)) + safe_num(row.get(COL_SALUD_VOL, 0))
        total_rebajas_llss = (
            safe_num(row.get(COL_AFP, 0))
            + safe_num(row.get(COL_CES_TRAB, 0))
            + safe_num(row.get(COL_APVI_MOD_B, 0))
            + safe_num(row.get(COL_TRAB_PESADO, 0))
            + min(salud_trab, tope_salud) if tope_salud > 0 else salud_trab
        )

        # ── Código de institución del trabajador ──
        cod_afp_inst    = safe_num(row.get(COL_AFP_INST, 0))
        cod_isapre_inst = safe_num(row.get(COL_ISAPRE_INST, 0))
        cod_mutual_inst = safe_num(row.get(COL_MUTUAL_INST, 0))
        cod_ccaf_inst   = safe_num(row.get(COL_CCAF_INST, 0))

        # ── Lookup instituciones ──
        id_afp_trab    = lookup(inst_afp,      "cod_lre", cod_afp_inst,    "id_afp",    "")
        id_isapre_trab = lookup(inst_salud,    "cod_lre", cod_isapre_inst, "id_inst",   "")
        id_mutual_trab = lookup(inst_mutuales, "cod_lre", cod_mutual_inst, "id_mutual", "")
        id_ccaf_trab   = lookup(inst_cajas,    "cod_lre", cod_ccaf_inst,   "id_ccaf",   "")
        cot_afp_trab   = lookup(inst_afp,      "cod_lre", cod_afp_inst,    "cot_afp",   0)

        # ── Cotización mutual (triangulación empleado → empresa) ──
        cot_mutual = 0.93
        if empresa_salida and not df_empresas.empty and "Empresa" in df_empresas.columns:
            emp2 = df_empresas[df_empresas["Empresa"] == empresa_salida]
            if not emp2.empty and "Cotización Mutual" in df_empresas.columns:
                cot_mutual = safe_num(emp2.iloc[0].get("Cotización Mutual", 0.93), 0.93)

        # ── Días reales del mes ──
        try:
            anio, mes = int(fecha_proceso[:4]), int(fecha_proceso[5:7])
            dias_reales_mes = calendar.monthrange(anio, mes)[1]
        except Exception:
            dias_reales_mes = 30

        CONCEPTOS_LICENCIA_COMPLETA = {
            "sueldoBase", "afp", "isapre", "cesEmpleado",
            "impuesto", "totalesEmpl", "mutual", "sis", "cesAporteCi"
        }

        # ── Generar fila por cada concepto ──
        for col_csv in cols_conceptos:
            id_concepto = equiv_map.get(col_csv, "")
            if not id_concepto:
                continue

            # Si licencia mes completo, solo incluir conceptos permitidos
            if dias_lic == dias_reales_mes and id_concepto not in CONCEPTOS_LICENCIA_COMPLETA:
                continue

            # Monto especial para isapre
            if id_concepto == "isapre":
                monto = monto_isapre
            else:
                monto = safe_num(row.get(col_csv, 0))

            # Saltar si monto es 0, excepto impuesto y cesEmpleado que siempre se incluyen
            CONCEPTOS_SIEMPRE = {"impuesto", "cesEmpleado"}
            if monto == 0 and id_concepto not in CONCEPTOS_SIEMPRE:
                continue

            # ── Id de institución ──
            id_institucion = ""
            if id_concepto in CONCEPTOS_ID_AFP:
                id_institucion = id_afp_trab
            elif id_concepto == "isapre":
                id_institucion = id_isapre_trab
            elif id_concepto == "mutual":
                id_institucion = id_mutual_trab
            elif id_concepto == "cajaCred":
                id_institucion = id_ccaf_trab

            # ── Afecto ──
            if id_concepto in CONCEPTOS_AFECTO_AFP:
                afecto = min(suma_haber_afecto, tope_afp) if tope_afp > 0 else suma_haber_afecto
            elif id_concepto in CONCEPTOS_AFECTO_CES:
                afecto = min(suma_haber_afecto, tope_ces) if tope_ces > 0 else suma_haber_afecto
            elif id_concepto == "totalesEmpl":
                afecto = suma_haber_afecto + suma_haber_exento
            elif id_concepto == "impuesto":
                afecto = suma_haber_afecto - total_rebajas_llss
            else:
                afecto = 0

            # ── Cotización de jubilación ──
            cot_jubilacion = ""
            if id_concepto == "cesEmpleado":
                cot_jubilacion = 0.6
            elif id_concepto == "afp":
                cot_jubilacion = cot_afp_trab
            elif id_concepto == "isapre" and id_institucion == "fonasa":
                cot_jubilacion = monto
            elif id_concepto == "mutual":
                cot_jubilacion = cot_mutual
            elif id_concepto == "sis":
                cot_jubilacion = safe_num(row_params.iloc[0].get("sis", 0)) if not row_params.empty else 0

            # ── Rentas no gravadas (solo si concepto = impuesto) ──
            rentas_no_grav = suma_haber_exento if id_concepto == "impuesto" else 0

            # ── Total rebajas LLSS (solo si concepto = impuesto) ──
            rebajas_llss = total_rebajas_llss if id_concepto == "impuesto" else 0

            filas.append({
                "Fecha de proceso":        fecha_proceso,
                "Id empleado":             rut,
                "Número de contrato":      numero_contrato,
                "Id del concepto":         id_concepto,
                "Monto del concepto":      monto,
                "Afecto":                  afecto,
                "Id de institución":       id_institucion,
                "Cotización de jubilación": cot_jubilacion,
                "Días de licencias":       dias_lic,
                "Días trabajados":         dias_trab,
                "Fecha de aplicación":     fecha_proceso,
                "Empresa":                 empresa_salida,
                "Total de rebajas por LLSS": rebajas_llss,
                "Rentas no gravadas":      rentas_no_grav,
                "Rebaja por zona extrema": rebaja_zona,
                "Jornada":                 "C",
                "Días de vacaciones":      dias_vac,
                "Monto Init":              monto_init,
                "Fase":                    1,
            })

        # ── Fila adicional: licenciaDias (si días de licencia > 0) ──
        if dias_lic > 0:
            filas.append({
                "Fecha de proceso":          fecha_proceso,
                "Id empleado":               rut,
                "Número de contrato":        numero_contrato,
                "Id del concepto":           "licenciaDias",
                "Monto del concepto":        dias_lic,
                "Afecto":                    0,
                "Id de institución":         "",
                "Cotización de jubilación":  "",
                "Días de licencias":         dias_lic,
                "Días trabajados":           dias_trab,
                "Fecha de aplicación":       fecha_proceso,
                "Empresa":                   empresa_salida,
                "Total de rebajas por LLSS": 0,
                "Rentas no gravadas":        0,
                "Rebaja por zona extrema":   rebaja_zona,
                "Jornada":                   "C",
                "Días de vacaciones":        dias_vac,
                "Monto Init":                monto_init,
                "Fase":                      1,
            })
    df_log_contratos = pd.DataFrame()
    if ruts_problema:
        df_log_contratos = construir_log_contratos(
            df_empleados, list(ruts_problema.keys()), fecha_proceso, ruts_problema
        )

    return pd.DataFrame(filas), df_log_contratos


# ─────────────────────────────────────────────
# GENERACIÓN EXCEL DE SALIDA
# ─────────────────────────────────────────────
def generar_excel_dt(df_salida):
    """Genera el Excel de salida con formato Rex+."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Liquidaciones"

    header_fill = PatternFill("solid", fgColor="1A2744")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    border = Border(
        bottom=Side(style="thin", color="E8EDF5"),
        right=Side(style="thin", color="E8EDF5")
    )

    cols = list(df_salida.columns)
    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = max(len(col) + 4, 14)

    for ri, row in enumerate(df_salida.itertuples(index=False), 2):
        fill = PatternFill("solid", fgColor="EAF0F8") if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.border = border
            cell.alignment = Alignment(vertical="center")

    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()


# ─────────────────────────────────────────────
# GENERACIÓN EXCEL LOG MÚLTIPLES CONTRATOS
# ─────────────────────────────────────────────
def generar_excel_log(df_log):
    """Genera el Excel de log de trabajadores con múltiples contratos."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Múltiples contratos"

    header_fill = PatternFill("solid", fgColor="C53030")
    header_font = Font(bold=True, color="FFFFFF", size=10)

    cols = list(df_log.columns)
    for ci, col in enumerate(cols, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[cell.column_letter].width = max(len(col) + 4, 18)

    for ri, row in enumerate(df_log.itertuples(index=False), 2):
        fill = PatternFill("solid", fgColor="FFF5F5") if ri % 2 == 0 else PatternFill("solid", fgColor="FFFFFF")
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.alignment = Alignment(vertical="center")

    ws.freeze_panes = "A2"
    wb.save(output)
    return output.getvalue()


# ─────────────────────────────────────────────
# COLUMNAS PARA VALIDACIONES (igual que Previred)
# ─────────────────────────────────────────────
COLS_HABERES_AFECTOS_DT = [
    "Sueldo(2101)", "Sobresueldo(2102)", "Comisiones(2103)", "Semana corrida(2104)",
    "Participación(2105)", "Gratificación(2106)", "Recargo 30% día domingo(2107)",
    "Remun. variable pagada en vacaciones(2108)", "Remun. variable pagada en clausura(2109)",
    "Aguinaldo(2110)", "Bonos u otras remun. fijas mensuales(2111)", "Tratos(2112)",
    "Bonos u otras remun. variables mensuales o superiores a un mes(2113)",
    "Ejercicio opción no pactada en contrato(2114)",
    "Beneficios en especie constitutivos de remun(2115)",
    "Remuneraciones bimestrales(2116)", "Remuneraciones trimestrales(2117)",
    "Remuneraciones cuatrimestral(2118)", "Remuneraciones semestrales(2119)",
    "Remuneraciones anuales(2120)", "Participación anual(2121)",
    "Gratificación anual(2122)", "Otras remuneraciones superiores a un mes(2123)",
    "Pago por horas de trabajo sindical(2124)", "Sueldo empresarial (2161)",
    "Subsidio por incapacidad laboral por licencia médica(2201)",
    "Beca de estudio(2202)", "Gratificaciones de zona(2203)"
]

COLS_HABERES_EXENTOS_DT = [
    "Otros ingresos no constitutivos de renta(2204)", "Colación(2301)",
    "Movilización(2302)", "Viáticos(2303)", "Asignación de pérdida de caja(2304)",
    "Asignación de desgaste herramienta(2305)", "Asignación familiar legal(2311)",
    "Gastos por causa del trabajo(2306)", "Gastos por cambio de residencia(2307)",
    "Sala cuna(2308)", "Asignación trabajo a distancia o teletrabajo(2309)",
    "Depósito convenido hasta UF 900(2347)", "Alojamiento por razones de trabajo(2310)",
    "Asignación de traslación(2312)", "Indemnización por feriado legal(2313)",
    "Indemnización años de servicio(2314)", "Indemnización sustitutiva del aviso previo(2315)",
    "Indemnización fuero maternal(2316)", "Pago indemnización a todo evento(2331)",
    "Indemnizaciones voluntarias tributables(2417)",
    "Indemnizaciones contractuales tributables(2418)"
]

COLS_DESCUENTOS_LEGALES_DT = [
    "Cotización obligatoria previsional (AFP o IPS)(3141)",
    "Cotización obligatoria salud 7%(3143)", "Cotización voluntaria para salud(3144)",
    "Cotización AFC - trabajador(3151)",
    "Cotizaciones técnico extranjero para seguridad social fuera de Chile(3146)",
    "Descuento depósito convenido hasta UF 900 anual(3147)",
    "Cotización APVi Mod A(3155)", "Cotización APVi Mod B hasta UF50(3156)",
    "Cotización APVc Mod A(3157)", "Cotización APVc Mod B hasta UF50(3158)",
    "Impuesto retenido por remuneraciones(3161)",
    "Impuesto retenido por indemnizaciones(3162)",
    "Mayor retención de impuestos solicitada por el trabajador(3163)",
    "Impuesto retenido por reliquidación remun. devengadas otros períodos(3164)",
    "Diferencia impuesto reliquidación remun. devengadas en este período(3165)",
    "Retención préstamo clase media 2020 (Ley 21.252) (3166)",
    "Rebaja zona extrema DL 889 (3167)"
]

COLS_OTROS_DESCUENTOS_DT = [
    "Cuota sindical 1(3171)", "Cuota sindical 2(3172)", "Cuota sindical 3(3173)",
    "Cuota sindical 4(3174)", "Cuota sindical 5(3175)", "Cuota sindical 6(3176)",
    "Cuota sindical 7(3177)", "Cuota sindical 8(3178)", "Cuota sindical 9(3179)",
    "Cuota sindical 10(3180)", "Crédito social CCAF(3110)",
    "Cuota vivienda o educación(3181)", "Crédito cooperativas de ahorro(3182)",
    "Otros descuentos autorizados y solicitados por el trabajador(3183)",
    "Cotización adicional trabajo pesado - trabajador(3154)",
    "Donaciones culturales y de reconstrucción(3184)", "Otros descuentos(3185)",
    "Pensiones de alimentos(3186)", "Descuento mujer casada(3187)",
    "Descuentos por anticipos y préstamos(3188)"
]

COLS_APORTES_EMPLEADOR_DT = [
    "AFC - Aporte empleador(4151)",
    "Aporte empleador seguro accidentes del trabajo y Ley SANNA(4152)",
    "Aporte adicional trabajo pesado - empleador(4154)",
    "Aporte empleador seguro invalidez y sobrevivencia(4155)",
    "APVC - Aporte Empleador(4157)"
]


def safe_sum_dt(df, cols):
    """Suma columnas que existen en el df."""
    cols_presentes = [c for c in cols if c in df.columns]
    if not cols_presentes:
        return pd.Series(0, index=df.index)
    return df[cols_presentes].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)


def validar_cuadraturas_dt(df, nombre_archivo):
    """Ejecuta las 6 validaciones de cuadratura para el archivo DT."""
    errores = []
    tol = 1

    df = df.copy()
    df["_hab_afectos"]    = safe_sum_dt(df, COLS_HABERES_AFECTOS_DT)
    df["_hab_exentos"]    = safe_sum_dt(df, COLS_HABERES_EXENTOS_DT)
    df["_desc_legales"]   = safe_sum_dt(df, COLS_DESCUENTOS_LEGALES_DT)
    df["_otros_desc"]     = safe_sum_dt(df, COLS_OTROS_DESCUENTOS_DT)
    df["_aportes_emp"]    = safe_sum_dt(df, COLS_APORTES_EMPLEADOR_DT)

    validaciones = [
        ("V1", "_hab_afectos",  "Total haberes imponibles y tributables(5210)",
         "Haberes afectos ≠ Total haberes imponibles y tributables(5210)"),
        ("V2", "_hab_exentos",  "Total haberes no imponibles y no tributables(5230)",
         "Haberes exentos ≠ Total haberes no imponibles y no tributables(5230)"),
        ("V3", "_desc_legales", "Total descuentos por cotizaciones del trabajador(5341)",
         "Descuentos legales ≠ Total descuentos por cotizaciones del trabajador(5341)"),
        ("V4", "_otros_desc",   "Total otros descuentos(5302)",
         "Otros descuentos ≠ Total otros descuentos(5302)"),
        ("V5", "_aportes_emp",  "Total aportes empleador(5410)",
         "Aportes empleador ≠ Total aportes empleador(5410)"),
    ]

    for codigo, col_calc, col_ctrl, mensaje in validaciones:
        if col_ctrl not in df.columns:
            continue
        ctrl = pd.to_numeric(df[col_ctrl], errors="coerce").fillna(0)
        calc = df[col_calc].fillna(0)
        mask = (calc - ctrl).abs() > tol
        for _, row in df[mask].iterrows():
            errores.append({
                "Archivo":          nombre_archivo,
                "RUT":              row.get(COL_RUT, "N/D"),
                "Validación":       codigo,
                "Descripción":      mensaje,
                "Valor calculado":  round(calc[row.name], 2),
                "Valor control":    round(ctrl[row.name], 2),
                "Diferencia":       round(calc[row.name] - ctrl[row.name], 2)
            })

    # V6: liquidez
    if "Total líquido(5501)" in df.columns:
        liq_calc = (df["_hab_afectos"] + df["_hab_exentos"]) - \
                   (df["_desc_legales"] + df["_otros_desc"])
        liq_ctrl = pd.to_numeric(df["Total líquido(5501)"], errors="coerce").fillna(0)
        mask = (liq_calc - liq_ctrl).abs() > tol
        for _, row in df[mask].iterrows():
            errores.append({
                "Archivo":          nombre_archivo,
                "RUT":              row.get(COL_RUT, "N/D"),
                "Validación":       "V6",
                "Descripción":      "(hab_afectos + hab_exentos) - (desc_legales + otros_desc) ≠ Total líquido(5501)",
                "Valor calculado":  round(liq_calc[row.name], 2),
                "Valor control":    round(liq_ctrl[row.name], 2),
                "Diferencia":       round(liq_calc[row.name] - liq_ctrl[row.name], 2)
            })

    return errores


# ─────────────────────────────────────────────
# INTERFAZ STREAMLIT DEL MÓDULO DT
# ─────────────────────────────────────────────
def render_modulo_dt(refs_compartidas):
    """
    Renderiza la interfaz del módulo DT dentro del tab de Streamlit.
    refs_compartidas: dict con archivos de referencia ya cargados por app_migracion.
    """
    st.markdown('<div class="section-title">🏛️ Migración Declaración Jurada DT</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Procesa el archivo descargado directamente desde la Dirección del Trabajo y genera el archivo de liquidaciones en detalle Rex+.</div>', unsafe_allow_html=True)

    # Pasos
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="step-card">
            <div class="step-label">PASO 1</div>
            <div class="step-title">Subir archivo DT</div>
            <div class="step-desc">Archivo CSV descargado desde el portal de la Dirección del Trabajo.</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="step-card">
            <div class="step-label">PASO 2</div>
            <div class="step-title">Subir listado de empleados</div>
            <div class="step-desc">Listado de empleados del período a procesar exportado desde Rex+.</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="step-card">
            <div class="step-label">PASO 3</div>
            <div class="step-title">Descargar resultado</div>
            <div class="step-desc">Se genera el archivo Excel de liquidaciones listo para importar.</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<hr class="rex-divider">', unsafe_allow_html=True)

    # ── Upload archivo DT ──
    st.markdown("### 📤 Archivo CSV de la Dirección del Trabajo")
    archivo_dt = st.file_uploader(
        "Selecciona el archivo CSV de la DT",
        type=["csv"],
        accept_multiple_files=False,
        key="dt_csv_upload",
        help="Archivo descargado desde el portal de la Dirección del Trabajo. El encabezado puede estar al inicio o al final."
    )

    # ── Upload listado empleados ──
    st.markdown("### 👥 Listado de empleados del período")
    archivo_empleados = st.file_uploader(
        "Sube el archivo listado_empleados.xlsx del período",
        type=["xlsx"],
        accept_multiple_files=False,
        key="dt_empleados_upload",
        help="Exportado desde Rex+. Debe contener columnas: Rut, Nombre, Empresa, Contrato."
    )

    # ── Upload listado empresas ──
    st.markdown("### 🏢 Listado de empresas del período")
    archivo_empresas = st.file_uploader(
        "Sube el archivo listado_empresas.xlsx del período",
        type=["xlsx"],
        accept_multiple_files=False,
        key="dt_empresas_upload",
        help="Exportado desde Rex+. Debe contener columnas: Empresa, Nombre, Cotización Mutual."
    )

    if archivo_empleados:
        st.markdown(f'<div class="alert-success">✅ Listado de empleados cargado: <b>{archivo_empleados.name}</b></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="alert-warning">⚠️ Debes subir el listado de empleados para ejecutar el proceso.</div>', unsafe_allow_html=True)

    if archivo_empresas:
        st.markdown(f'<div class="alert-success">✅ Listado de empresas cargado: <b>{archivo_empresas.name}</b></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="alert-warning">⚠️ Debes subir el listado de empresas para ejecutar el proceso.</div>', unsafe_allow_html=True)

    # ── Upload parámetros mensuales ──
    st.markdown("### 📅 Parámetros mensuales")
    archivo_params_dt = st.file_uploader(
        "Sube el archivo parametrosMesuales.xlsx del período",
        type=["xlsx"],
        accept_multiple_files=False,
        key="dt_params_upload",
        help="Archivo con los parámetros legales del mes a procesar."
    )
    if archivo_params_dt:
        st.markdown(f'<div class="alert-success">✅ Parámetros mensuales cargados: <b>{archivo_params_dt.name}</b></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="alert-warning">⚠️ Debes subir el archivo de parámetros mensuales para ejecutar el proceso.</div>', unsafe_allow_html=True)

    if not archivo_dt or not archivo_empleados or not archivo_empresas or not archivo_params_dt:
        return

    st.markdown(f'<div class="alert-success">✅ Archivo DT cargado: <b>{archivo_dt.name}</b></div>', unsafe_allow_html=True)

    # ── Determinar fecha de proceso ──
    fecha_proceso, fecha_ok = extraer_fecha_dt(archivo_dt.name)

    if not fecha_ok:
        st.markdown("""
        <div class="alert-warning">
            ⚠️ <b>No se pudo determinar la fecha de proceso</b> desde el nombre del archivo.<br>
            Por favor selecciona el mes y año correspondiente.
        </div>""", unsafe_allow_html=True)

        col_m, col_a, _ = st.columns([1, 1, 3])
        with col_m:
            mes_sel = st.selectbox("Mes", options=list(range(1, 13)),
                                   format_func=lambda x: f"{x:02d} - {['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'][x-1]}",
                                   key="dt_mes_sel")
        with col_a:
            anio_actual = datetime.now().year
            anio_sel = st.selectbox("Año", options=list(range(anio_actual - 5, anio_actual + 2)),
                                    index=5, key="dt_anio_sel")
        fecha_proceso = f"{anio_sel}-{mes_sel:02d}"

    st.markdown(f'<div class="alert-success">📅 Fecha de proceso: <b>{fecha_proceso}</b></div>', unsafe_allow_html=True)

    # ── Botón ejecutar ──
    if st.button("▶ Ejecutar proceso DT", key="dt_btn_ejecutar"):
        with st.spinner("Procesando archivo..."):
            try:
                # Leer CSV DT
                df_dt = leer_csv_dt(archivo_dt)

                # Leer empleados
                df_empleados = cargar_empleados(archivo_empleados)

                # Leer empresas
                df_empresas_periodo = pd.read_excel(archivo_empresas, header=1)
                df_empresas_periodo.columns = [str(c).strip() for c in df_empresas_periodo.columns]
                if "Nombre" in df_empresas_periodo.columns:
                    df_empresas_periodo["Nombre"] = df_empresas_periodo["Nombre"].astype(str).str.strip()
                if "Empresa" in df_empresas_periodo.columns:
                    df_empresas_periodo["Empresa"] = df_empresas_periodo["Empresa"].astype(str).str.strip()

                # Leer parámetros mensuales
                df_params_periodo = pd.read_excel(archivo_params_dt)

                # Verificar que el mes existe en los parámetros
                if "mes_Proc" in df_params_periodo.columns:
                    df_params_periodo["mes_Proc"] = df_params_periodo["mes_Proc"].astype(str).str.strip()
                    if fecha_proceso not in df_params_periodo["mes_Proc"].values:
                        st.markdown(f'<div class="alert-error">❌ <b>No se encontraron parámetros para {fecha_proceso}</b> en el archivo subido.</div>', unsafe_allow_html=True)
                        st.stop()

                # Inyectar en refs
                refs_dt = dict(refs_compartidas)
                refs_dt["listado_empresas"] = df_empresas_periodo
                refs_dt["parametros"] = df_params_periodo

                # Ejecutar validaciones de cuadratura
                errores_val = validar_cuadraturas_dt(df_dt, archivo_dt.name)

            except Exception as e:
                st.markdown(f'<div class="alert-error">❌ Error al leer los archivos: <b>{e}</b></div>', unsafe_allow_html=True)
                return

        st.markdown('<hr class="rex-divider">', unsafe_allow_html=True)
        st.markdown("### 🔍 Resultado de validaciones")

        # ── Mostrar errores de validación si hay ──
        if errores_val:
            st.markdown(f"""
            <div class="alert-error">
                ❌ <b>No se puede generar el archivo de salida.</b><br>
                Se encontraron <b>{len(errores_val)} error(es)</b> de validación en los registros procesados.
            </div>""", unsafe_allow_html=True)

            with st.expander("📋 Ver log de errores detallado"):
                df_errores = pd.DataFrame(errores_val)
                st.dataframe(df_errores, use_container_width=True, hide_index=True)
                csv_log = df_errores.to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="⬇️ Descargar log de errores (.csv)",
                    data=csv_log,
                    file_name=f"log_errores_dt_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="dt_btn_log_errores"
                )
            return

        st.markdown("""
        <div class="alert-success">
            ✅ <b>Todas las validaciones se cumplieron correctamente.</b><br>
            Los registros cuadran sin diferencias.
        </div>""", unsafe_allow_html=True)

        # ── Generar archivo de salida ──
        with st.spinner("Generando archivo de salida..."):
            try:
                df_salida, df_log_contratos = generar_filas_dt(df_dt, fecha_proceso, refs_dt, df_empleados, df_empresas_externo=df_empresas_periodo)
            except Exception as e:
                st.markdown(f'<div class="alert-error">❌ Error al generar el archivo de salida: <b>{e}</b></div>', unsafe_allow_html=True)
                return

        # ── Mostrar log de problemas de contrato si hay ──
        if not df_log_contratos.empty:
            st.markdown(f"""
            <div class="alert-warning">
                ⚠️ <b>{len(df_log_contratos["Rut"].unique())} trabajador(es)</b> fueron excluidos por problemas en la resolución de contrato.<br>
                Descarga el log para revisarlos.
            </div>""", unsafe_allow_html=True)

            with st.expander(f"👁️ Ver trabajadores excluidos por contrato ({len(df_log_contratos['Rut'].unique())})"):
                st.dataframe(df_log_contratos, use_container_width=True, hide_index=True)

            log_cont_bytes = generar_excel_log(df_log_contratos)
            st.download_button(
                label="⬇️ Descargar log_multiples_contratos.xlsx",
                data=log_cont_bytes,
                file_name=f"log_multiples_contratos_{fecha_proceso}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dt_btn_log_contratos"
            )

        if df_salida.empty:
            st.markdown('<div class="alert-warning">⚠️ No se generaron registros. Verifica que el archivo y los parámetros sean correctos.</div>', unsafe_allow_html=True)
            return

        st.markdown(f"""
        <div class="alert-success">
            ✅ <b>Archivo generado exitosamente.</b><br>
            Se procesaron <b>{df_salida["Id empleado"].nunique()} trabajadores</b> y se generaron <b>{len(df_salida)} registros</b>.
        </div>""", unsafe_allow_html=True)

        with st.expander("👁️ Vista previa del archivo de salida"):
            st.dataframe(df_salida.head(50), use_container_width=True, hide_index=True)

        excel_bytes = generar_excel_dt(df_salida)
        st.download_button(
            label="⬇️ Descargar archivo de salida (.xlsx)",
            data=excel_bytes,
            file_name=f"liquidaciones_dt_{fecha_proceso}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dt_btn_descarga"
        )
