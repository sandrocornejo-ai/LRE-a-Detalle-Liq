"""
migracion_lre_dt.py
Migración desde LRE DT — Etapa 1: Armado del maestro.

1. Pide todos los CSV del Libro de Remuneraciones Electrónico descargados desde la DT.
2. Verifica que el nombre de CADA archivo incluya año-mes; si alguno no cumple,
   aborta todo el proceso ("Nombre de archivo no cumple los requisitos").
   Todos los archivos deben ser del mismo año; si no, aborta ("Los archivos son de años distintos").
3. En cada archivo busca el encabezado y, si no está arriba, lo deja como primera fila.
4. Consolida todos los archivos en uno solo con "Mes de proceso" (yyyy-mm) como primera columna.
   Agrega MpRut_entrada (Mes de proceso + Rut) después de Rut trabajador(1101), en verde claro.
5. Descarga el consolidado en Excel.

6. El usuario elige el rango de meses corridos a procesar (uno, varios o todos). Los cálculos usan
   todos los meses subidos como historia; el maestro solo trae el rango elegido.

Ejecutar:  streamlit run migracion_lre_dt.py
"""

import io
import os
import re
import traceback
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
MSG_NOMBRE_INVALIDO = "Nombre de archivo no cumple los requisitos"
MSG_ANIOS_DISTINTOS = "Los archivos son de años distintos"
COL_MES_PROCESO = "Mes de proceso"
COL_RUT = "Rut trabajador(1101)"
COL_MP_RUT = "MpRut_entrada"            # Mes de proceso + Rut trabajador (concatenados)

# Columnas agregadas por el programa (no vienen de la DT) → se pintan verde claro en el Excel
COL_AFP_DT = "(1141)"                   # código AFP en el LRE → columna 'AFP(1141)'
COL_NOMBRE_AFP = "nombreafp_entrada"    # afp_id_rex buscado en Instituciones.xlsx, hoja AFP

COL_PORC_AFP = "porcafp_entrada"        # cot_hist_afp de cot_afp_hist.xlsx
COL_SIS = "Sis_entrada"                 # sis_hist de cot_afp_hist.xlsx

COL_SALUD_DT = "(1143)"                 # código salud en el LRE → 'FONASA - ISAPRE(1143)'
COL_NOMBRE_SALUD = "nombresalud_entrada"  # salud_id_rex de Instituciones.xlsx, hoja SALUD
COL_CAJA_DT = "(1110)"                  # código CCAF en el LRE → 'CCAF(1110)'
COL_NOMBRE_CAJA = "nombrecaja_entrada"  # caja_id_rex de Instituciones.xlsx, hoja CAJAS
COL_MUTUAL_DT = "(1152)"                # código mutual en el LRE → 'Org. administrador ley 16.744(1152)'
COL_NOMBRE_MUTUAL = "nombreMUTUALa_entrada"  # mutual_id_rex de Instituciones.xlsx, hoja MUTUALES
COL_PORC_MUTUAL = "Porcmutual_entrada"  # 'Cotización Mutual' de listado_empresas, buscando Empresa_entrada

# Columnas desde Parámetros mensuales (buscando por Mes de proceso = mes_Proc); van después de (5565), en este orden
COL_ULTIMA_LRE = "(5565)"               # 'Total indemnizaciones no tributables(5565)'
PARAMETROS_A_TRAER = [                  # (columna en parametrosMesuales, columna nueva)
    ("topeImp_pesos_afp",   "topeimp_entrada"),
    ("topeCes_pesos",       "topeces_entrada"),
    ("topeSalud_pesos",     "topesalud_entrada"),
    ("aporte_Ccaf",         "porcAportecaja_entrada"),
    ("Aporte AFP",          "aporteAfp_entrada"),
    ("Seg Social Exp vida", "aporteExpvida_entrada"),
    ("aporteFAPPBAC",       "aporteFappbac_entrada"),
]
COL_REBAJA_LLSS = "rebajallss_entrada"

# Último mes imponible sin licencia (va después de 'Nro días de licencia médica en el mes(1116)')
COL_DIAS_LIC_DT = "(1116)"
COL_IMP_TRIB_DT = "(5210)"              # 'Total haberes imponibles y tributables(5210)'
COL_ULT_IMP_SIN_LIC = "ultImpSinLic_entrada"
COL_IMP_MES_ANT_LIC = "impoMesAntporDdeLic_entrada"   # (ultImpSinLic_entrada / 30) × días licencia(1116); va después de ultImpSinLic_entrada
COL_IMPONIBLE_LIC = "imponibleLic_entrada"   # si 1116 > 0: impoMesAntporDdeLic_entrada + 5210 + 5220 (con tope); si no, 0
COL_IMP_NO_TRIB_DT = "(5220)"           # 'Total haberes imponibles no tributables(5220)'
TXT_IMP_NO_ENCONTRADO = "imponible no encontrado"
AMARILLO = "FFFF00"

# Sueldo de contrato (va después de 'Tasa indemnización a todo evento(1132)')
COL_TASA_IND_DT = "(1132)"
COL_SUELDO_DT = "(2101)"
COL_DIAS_TRAB_DT = "(1115)"
COL_SUELDO_CONTRATO = "sueldoContrato_entrada"
TXT_SUELDO_NO_ENCONTRADO = "sueldo no encontrado"

# Datos que pueden quedar "no encontrados" y completarse a mano en la ventana emergente
DATOS_MANUALES = {
    # pedir=False: no se pide en la ventana; se resuelve al ingresar el sueldo de contrato del mismo Rut
    COL_ULT_IMP_SIN_LIC: {"texto": TXT_IMP_NO_ENCONTRADO, "etiqueta": "Último imponible sin licencia", "pedir": False},
    COL_SUELDO_CONTRATO: {"texto": TXT_SUELDO_NO_ENCONTRADO, "etiqueta": "Sueldo de contrato", "pedir": True},
}
NARANJO_CLARO = "FCE4D6"                # celda ultImpSinLic_entrada ingresada manualmente
CLAVE_IMP_MANUAL = "imponibles_manuales"  # st.session_state: {rut normalizado: imponible}
CODIGOS_REBAJA_LLSS = ["(3141)", "(3143)", "(3151)", "(3156)", "(3158)", "(3167)", "(3154)"]

# Búsquedas simples en Instituciones.xlsx: (columna LRE, hoja, columna clave, campo a traer, columna nueva)
BUSQUEDAS_INSTITUCIONES = [
    (COL_SALUD_DT, "SALUD", "salud_cod_dt", "salud_id_rex", COL_NOMBRE_SALUD),
    (COL_CAJA_DT,  "CAJAS", "caja_cod_dt",  "caja_id_rex",  COL_NOMBRE_CAJA),
    (COL_MUTUAL_DT, "MUTUALES", "mutual_cod_dt", "mutual_id_rex", COL_NOMBRE_MUTUAL),
]

COL_FECHA_INI_DT = "(1102)"             # 'Fecha inicio contrato(1102)'
COL_EMPRESA = "Empresa_entrada"         # listado_empleados: columna BG (Empresa)
COL_NUM_CONTRATO = "Numcontrato_entrada"  # listado_empleados: columna BE (Contrato)
COL_HORAS_SEM = "horasSema_entrada"     # listado_empleados: Horas Semanales (col. CE), misma búsqueda Rut + Fecha inicio
COL_ID_EMPRESA = "idEmpresa_entrada"    # listado_empresas: Empresa (código), buscando Empresa_entrada en Nombre

COLUMNAS_NUEVAS = [COL_MP_RUT, COL_ID_EMPRESA, COL_EMPRESA, COL_NUM_CONTRATO, COL_HORAS_SEM, COL_NOMBRE_AFP, COL_PORC_AFP, COL_SIS,
                   COL_NOMBRE_SALUD, COL_NOMBRE_CAJA, COL_NOMBRE_MUTUAL, COL_PORC_MUTUAL] \
                  + [nueva for _, nueva in PARAMETROS_A_TRAER] + [COL_REBAJA_LLSS, COL_ULT_IMP_SIN_LIC, COL_IMP_MES_ANT_LIC, COL_IMPONIBLE_LIC, COL_SUELDO_CONTRATO]
COLUMNAS_TASA = [COL_PORC_AFP, COL_SIS, COL_PORC_MUTUAL]  # en porcentaje (x 100), 2 decimales en el Excel
VERDE_CLARO = "C6EFCE"
MARCA_ENCABEZADO = "Rut trabajador"      # texto que identifica la fila de encabezado del LRE
SEPARADOR = ";"

MESES = {
    "enero": "01", "ene": "01",
    "febrero": "02", "feb": "02",
    "marzo": "03", "mar": "03",
    "abril": "04", "abr": "04",
    "mayo": "05", "may": "05",
    "junio": "06", "jun": "06",
    "julio": "07", "jul": "07",
    "agosto": "08", "ago": "08",
    "septiembre": "09", "setiembre": "09", "sept": "09", "sep": "09", "set": "09",
    "octubre": "10", "oct": "10",
    "noviembre": "11", "nov": "11",
    "diciembre": "12", "dic": "12",
}
_RE_MES = "|".join(sorted(MESES, key=len, reverse=True))   # más largas primero


# ─────────────────────────────────────────────
# TABLA DE INSTITUCIONES (data/Instituciones.xlsx)
# ─────────────────────────────────────────────
DIR_DATA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
RUTA_INSTITUCIONES = os.path.join(DIR_DATA, "Instituciones.xlsx")
COLS_COT_AFP_HIST = ["id_afp_hist", "cot_hist_afp", "sis_hist"]
COLS_PARAMETROS = ["mes_Proc"] + [col for col, _ in PARAMETROS_A_TRAER]

# hoja → (columnas obligatorias, columna código DT)
HOJAS_INSTITUCIONES = {
    "AFP":      (["afp_cod_dt", "afp_id_rex", "afp_nombre", "afp_rut", "afp_codPrev", "afp_cot"], "afp_cod_dt"),
    "SALUD":    (["salud_cod_dt", "salud_id_rex", "salud_nombre", "salud_rut", "salud_cod_prev"], "salud_cod_dt"),
    "CAJAS":    (["caja_cod_dt", "caja_id_rex", "caja_nombre", "caja_rut", "caja_cod_prev"], "caja_cod_dt"),
    "MUTUALES": (["mutual_cod_dt", "mutual_id_rex", "mutual_nombre", "mutual_rut", "mutual_cod_prev"], "mutual_cod_dt"),
}


class ErrorArchivo(Exception):
    """Error de estructura en un archivo DT."""


def _cod(valor):
    """Normaliza un código a texto sin decimales ni espacios: 31 / 31.0 / ' 31 ' → '31'."""
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    return txt[:-2] if txt.endswith(".0") else txt


def cargar_instituciones(ruta=RUTA_INSTITUCIONES):
    """
    Lee Instituciones.xlsx (hojas AFP, SALUD, CAJAS, MUTUALES).
    Todo se lee como texto para conservar ceros a la izquierda (códigos Previred '03', '08').
    Solo afp_cot se convierte a número (tasa %).
    Retorna dict {hoja: DataFrame}. Lanza ErrorArchivo si falta el archivo, una hoja o una columna.
    """
    if not os.path.exists(ruta):
        raise ErrorArchivo(f"No se encontró el archivo de instituciones: {ruta}")
    hojas = pd.read_excel(ruta, sheet_name=None, dtype=str)
    tablas = {}
    for hoja, (cols_req, col_cod) in HOJAS_INSTITUCIONES.items():
        if hoja not in hojas:
            raise ErrorArchivo(f"Instituciones.xlsx no tiene la hoja '{hoja}'.")
        df = hojas[hoja]
        df.columns = [str(c).strip() for c in df.columns]
        faltan = [c for c in cols_req if c not in df.columns]
        if faltan:
            raise ErrorArchivo(f"Instituciones.xlsx, hoja '{hoja}': faltan columnas {faltan}.")
        df = df.dropna(how="all").fillna("")
        for c in df.columns:
            df[c] = df[c].astype(str).str.strip()
        df[col_cod] = df[col_cod].map(_cod)
        if "afp_cot" in df.columns:
            df["afp_cot"] = pd.to_numeric(df["afp_cot"].str.replace(",", ".", regex=False), errors="coerce")
        tablas[hoja] = df.reset_index(drop=True)
    return tablas


def cargar_cot_afp_hist(archivo):
    """
    Lee el archivo subido de cotizaciones históricas AFP y SIS (ej. cot_afp_hist.xlsx).
    Clave id_afp_hist = mes + id AFP (ej. '2026-01capital'). cot_hist_afp y sis_hist como número.
    """
    archivo.seek(0)
    df = pd.read_excel(archivo, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    faltan = [c for c in COLS_COT_AFP_HIST if c not in df.columns]
    if faltan:
        raise ErrorArchivo(f"{archivo.name}: faltan columnas {faltan}.")
    df = df.dropna(subset=["id_afp_hist"]).copy()
    df["id_afp_hist"] = df["id_afp_hist"].astype(str).str.strip().str.lower()
    for c in ("cot_hist_afp", "sis_hist"):
        df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    duplicados = df.loc[df["id_afp_hist"].duplicated(), "id_afp_hist"].unique().tolist()
    if duplicados:
        raise ErrorArchivo(f"{archivo.name}: id_afp_hist duplicados {duplicados[:10]}.")
    return df.reset_index(drop=True)


def _leer_excel_con_encabezado(archivo, columnas_clave, max_filas=10):
    """
    Lee un Excel exportado desde Rex+ detectando la fila de encabezado
    (la primera, dentro de las primeras 'max_filas', que contiene todas las columnas_clave).
    Así funciona tanto si el archivo trae un título arriba como si no.
    """
    archivo.seek(0)
    crudo = pd.read_excel(archivo, header=None, dtype=object)
    claves = [c.lower() for c in columnas_clave]
    for i in range(min(max_filas, len(crudo))):
        fila = [str(v).strip().lower() for v in crudo.iloc[i].tolist()]
        if all(c in fila for c in claves):
            df = crudo.iloc[i + 1:].reset_index(drop=True)
            df.columns = _nombres_unicos(crudo.iloc[i].tolist())
            return df.dropna(how="all")
    raise ErrorArchivo(f"{archivo.name}: no se encontró la fila de encabezado con {columnas_clave}.")


def _nombres_unicos(nombres):
    """Evita columnas con el mismo nombre (Rex+ repite, p. ej., 'Empresa'): 'Empresa', 'Empresa (2)', ..."""
    vistos, salida = {}, []
    for n in nombres:
        n = "" if pd.isna(n) else str(n).strip()
        if n in vistos:
            vistos[n] += 1
            salida.append(f"{n} ({vistos[n]})")
        else:
            vistos[n] = 1
            salida.append(n)
    return salida


def _indice_letra(letra):
    """'A' → 0, 'BE' → 56, 'BG' → 58."""
    idx = 0
    for ch in letra.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def _col_por_letra(df, letra, nombre_esperado):
    """Columna por letra de Excel. Avisa (en attrs) si el encabezado no se parece al esperado."""
    idx = _indice_letra(letra)
    if idx >= len(df.columns):
        raise ErrorArchivo(f"El archivo no tiene columna {letra} ({nombre_esperado}); tiene {len(df.columns)} columnas.")
    return df.columns[idx]


def _col_por_nombre_o_letra(df, nombre, letra):
    """Columna por nombre (sin distinguir mayúsculas); si no existe, por letra de Excel (ej. 'BG')."""
    for c in df.columns:
        if str(c).strip().lower() == nombre.lower():
            return c
    idx = 0
    for ch in letra.upper():
        idx = idx * 26 + (ord(ch) - 64)
    if idx - 1 < len(df.columns):
        return df.columns[idx - 1]
    raise ErrorArchivo(f"No se encontró la columna '{nombre}' (columna {letra}).")


def normalizar_rut(valor):
    """'12.458.987-8' / ' 12458987-8 ' / '9271991-k' → '12458987-8' / '9271991-K'."""
    if pd.isna(valor):
        return ""
    return str(valor).replace(".", "").replace(" ", "").strip().upper()


def normalizar_fecha(valor):
    """Fecha a texto dd/mm/aaaa. Acepta fecha de Excel, '04/03/2021', '04-03-2021' o '2021-03-04'."""
    if pd.isna(valor) or str(valor).strip() == "":
        return ""
    if isinstance(valor, (datetime, pd.Timestamp)):
        return valor.strftime("%d/%m/%Y")
    txt = str(valor).strip()
    formato_iso = re.fullmatch(r"\d{4}-\d{2}-\d{2}( .*)?", txt) is not None
    fecha = pd.to_datetime(txt, dayfirst=not formato_iso, errors="coerce")
    return fecha.strftime("%d/%m/%Y") if pd.notna(fecha) else txt


def cargar_empleados(archivo):
    """
    Lee listado_empleados.xlsx (Rex+). Devuelve DataFrame con:
      clave (Rut + Fecha inicio contrato, normalizados), Empresa (col. BG), Contrato (col. BE).
    """
    df = _leer_excel_con_encabezado(archivo, ["Rut"])
    col_rut = _col_por_nombre_o_letra(df, "Rut", "A")
    col_ini = next((c for c in df.columns if "inicio" in str(c).lower() and "contrato" in str(c).lower()), None)
    if col_ini is None:
        raise ErrorArchivo(f"{archivo.name}: no se encontró la columna 'Fecha inicio contrato'.")
    col_emp = _col_por_letra(df, "BG", "Empresa")    # posición fija indicada: BG
    col_con = _col_por_letra(df, "BE", "Contrato")   # posición fija indicada: BE
    col_horas = _col_por_nombre_o_letra(df, "Horas Semanales", "CE")
    out = pd.DataFrame({
        "clave": df[col_rut].map(normalizar_rut) + df[col_ini].map(normalizar_fecha),
        "Empresa": df[col_emp],
        "Contrato": df[col_con],
        "Horas Semanales": df[col_horas],
    })
    out = out[out["clave"] != ""].reset_index(drop=True)
    out.attrs["columnas_usadas"] = {"Rut": col_rut, "Fecha inicio": col_ini,
                                    "Empresa (BG)": col_emp, "Contrato (BE)": col_con,
                                    "Horas Semanales": col_horas}
    return out


def cargar_parametros(archivo):
    """
    Lee parametrosMesuales.xlsx (una fila por mes, clave mes_Proc = 'yyyy-mm').
    Por ahora solo se carga y valida.
    """
    archivo.seek(0)
    df = pd.read_excel(archivo, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    faltan = [c for c in COLS_PARAMETROS if c not in df.columns]
    if faltan:
        raise ErrorArchivo(f"{archivo.name}: faltan columnas {faltan}.")
    df = df.dropna(subset=["mes_Proc"]).copy()
    df["mes_Proc"] = df["mes_Proc"].astype(str).str.strip().str[:7]   # '2026-01-01 00:00:00' → '2026-01'
    for col, _ in PARAMETROS_A_TRAER:
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", ".", regex=False), errors="coerce")
    duplicados = df.loc[df["mes_Proc"].duplicated(), "mes_Proc"].unique().tolist()
    if duplicados:
        raise ErrorArchivo(f"{archivo.name}: meses repetidos en mes_Proc {duplicados}.")
    return df.reset_index(drop=True)


def cargar_empresas(archivo):
    """Lee listado_empresas.xlsx (Rex+). Por ahora solo se carga y valida."""
    df = _leer_excel_con_encabezado(archivo, ["Empresa", "Nombre"])
    if df.empty:
        raise ErrorArchivo(f"{archivo.name}: no tiene registros.")
    if "Cotización Mutual" not in df.columns:
        raise ErrorArchivo(f"{archivo.name}: falta la columna 'Cotización Mutual'.")
    df["Cotización Mutual"] = pd.to_numeric(df["Cotización Mutual"].astype(str).str.replace(",", ".", regex=False),
                                            errors="coerce")
    return df.reset_index(drop=True)


def buscar_institucion(tablas, hoja, cod_dt, campo):
    """
    Devuelve el valor de 'campo' para el código DT dado en la hoja indicada, o "" si no existe.
    Ej: buscar_institucion(inst, "AFP", 31, "afp_id_rex") → 'capital'
    """
    df = tablas.get(hoja)
    if df is None or campo not in df.columns:
        return ""
    col_cod = HOJAS_INSTITUCIONES[hoja][1]
    fila = df[df[col_cod] == _cod(cod_dt)]
    return fila.iloc[0][campo] if not fila.empty else ""


# ─────────────────────────────────────────────
# 1. PERÍODO DESDE EL NOMBRE DEL ARCHIVO
# ─────────────────────────────────────────────
def _anio4(txt):
    return f"20{txt}" if len(txt) == 2 else txt


def extraer_periodo(nombre_archivo):
    """
    Devuelve 'yyyy-mm' si el nombre del archivo incluye año-mes, o None.

    Formatos aceptados (sin distinguir mayúsculas; separador _ - espacio o punto):
      Ene_26_declaracion.csv · Julio_2026.csv · 2026_enero.csv · 26-ene.csv
      2026-01.csv · 2026_01.csv · 202601.csv
    """
    nombre = os.path.splitext(os.path.basename(str(nombre_archivo)))[0].lower()

    # Mes + año   (ene_26, julio_2026)
    m = re.search(rf"(?<![a-z])({_RE_MES})[_\-\s\.]*(\d{{4}}|\d{{2}})(?!\d)", nombre)
    if m:
        return f"{_anio4(m.group(2))}-{MESES[m.group(1)]}"

    # Año + mes   (2026_enero, 26-ene)
    m = re.search(rf"(?<!\d)(\d{{4}}|\d{{2}})[_\-\s\.]*({_RE_MES})(?![a-z])", nombre)
    if m:
        return f"{_anio4(m.group(1))}-{MESES[m.group(2)]}"

    # Numérico    (2026-01, 2026_01, 202601) — no debe ser parte de un número más largo
    m = re.search(r"(?<!\d)(20\d{2})[_\-\.]?(0[1-9]|1[0-2])(?!\d)", nombre)
    if m:
        return f"{m.group(1)}-{m.group(2)}"

    return None


def validar_nombres(archivos):
    """Retorna (validos, invalidos): validos = [(archivo, 'yyyy-mm')], invalidos = [nombre]."""
    validos, invalidos = [], []
    for f in archivos:
        periodo = extraer_periodo(f.name)
        if periodo:
            validos.append((f, periodo))
        else:
            invalidos.append(f.name)
    return validos, invalidos


# ─────────────────────────────────────────────
# 2. LECTURA DE CADA ARCHIVO (encabezado como primera fila)
# ─────────────────────────────────────────────
def _decodificar(contenido):
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return contenido.decode(enc)
        except UnicodeDecodeError:
            continue
    raise ErrorArchivo("No se pudo leer la codificación del archivo.")


def _buscar_col_rut(df):
    """Nombre real de la columna Rut (1101), tolerando espacios: 'Rut trabajador(1101)' / 'Rut trabajador (1101)'."""
    return next((c for c in df.columns if "(1101)" in str(c)), None)


def _a_numero(serie):
    """Convierte a número solo los valores que son números puros (enteros o con coma decimal).
    RUT, fechas y textos se mantienen como texto."""
    s = serie.fillna("").astype(str).str.strip()
    no_vacios = s[s != ""]
    if no_vacios.empty or not no_vacios.str.fullmatch(r"-?\d+(,\d+)?").all():
        return serie
    if no_vacios.str.contains(",").any():
        return pd.to_numeric(s.str.replace(",", ".", regex=False), errors="coerce")
    return pd.to_numeric(s, errors="coerce").astype("Int64")


def leer_archivo_dt(archivo):
    """
    Lee un CSV de la DT. Busca la fila de encabezado (la que contiene 'Rut trabajador'),
    la deja como primera fila y el resto como datos (sin líneas vacías).
    Retorna (DataFrame, posicion_encabezado) — posición 1-based en el archivo original.
    """
    archivo.seek(0)
    lineas = _decodificar(archivo.read()).splitlines()

    idx_enc = next((i for i, l in enumerate(lineas) if MARCA_ENCABEZADO in l), None)
    if idx_enc is None:
        raise ErrorArchivo(f"No se encontró la fila de encabezado ('{MARCA_ENCABEZADO}...').")

    datos = [l for i, l in enumerate(lineas) if i != idx_enc and l.strip()]
    contenido = "\n".join([lineas[idx_enc]] + datos)
    df = pd.read_csv(io.StringIO(contenido), sep=SEPARADOR, dtype=str, keep_default_na=False)

    for col in df.columns:
        df[col] = _a_numero(df[col])
    return df, idx_enc + 1


# ─────────────────────────────────────────────
# 3. CONSOLIDACIÓN
# ─────────────────────────────────────────────
def consolidar(validos, instituciones, cot_afp_hist, empleados, empresas, parametros):
    """
    Une todos los archivos en un DataFrame con 'Mes de proceso' como primera columna,
    ordenado por mes. Retorna (df_consolidado, resumen).
    """
    partes, resumen, descartadas = [], [], []
    for f, periodo in sorted(validos, key=lambda x: (x[1], x[0].name)):
        try:
            df, pos_enc = leer_archivo_dt(f)
            # Si el CSV ya trae columnas con el nombre de las que genera el programa, se descartan
            repetidas = [c for c in [COL_MES_PROCESO] + COLUMNAS_NUEVAS if c in df.columns]
            if repetidas:
                df = df.drop(columns=repetidas)
                descartadas.append((f.name, repetidas))
            df.insert(0, COL_MES_PROCESO, periodo)
        except Exception as e:  # noqa: BLE001
            e.archivo_origen = f.name   # para indicar en el log qué archivo falló
            raise
        partes.append(df)
        resumen.append({
            "Archivo": f.name,
            COL_MES_PROCESO: periodo,
            "Registros": len(df),
            "Encabezado estaba en": "primera fila" if pos_enc == 1 else f"fila {pos_enc} (movido)",
        })
    df_cons = pd.concat(partes, ignore_index=True, sort=False)
    df_cons = agregar_columnas_nuevas(df_cons, instituciones, cot_afp_hist, empleados, empresas, parametros)
    df_cons.attrs["columnas_descartadas"] = descartadas
    return df_cons, resumen


def _buscar_col_codigo(df, codigo):
    """Nombre real de la columna que contiene '(codigo)', ej. '(1141)' → 'AFP(1141)'."""
    return next((c for c in df.columns if codigo in str(c)), None)


def agregar_columnas_nuevas(df, instituciones, cot_afp_hist, empleados, empresas, parametros):
    """
    Agrega las columnas calculadas por el programa (se pintan verde claro en el Excel):
      - MpRut_entrada    : Mes de proceso (col. A) + Rut trabajador; va después de Rut trabajador(1101).
      - nombreafp_entrada: afp_id_rex de Instituciones.xlsx (hoja AFP) buscando AFP(1141) en afp_cod_dt;
                           va después de AFP(1141). Queda vacío si el código no existe en la tabla.
      - porcafp_entrada / Sis_entrada: cot_hist_afp y sis_hist (x 100) de cot_afp_hist.xlsx, buscando por
                           id_afp_hist = Mes de proceso + nombreafp_entrada; van después de nombreafp_entrada.
                           Quedan vacías si la clave no existe.
      - nombresalud_entrada: salud_id_rex (hoja SALUD, por salud_cod_dt); va después de FONASA - ISAPRE(1143).
      - nombrecaja_entrada : caja_id_rex  (hoja CAJAS, por caja_cod_dt);  va después de CCAF(1110).
      - nombreMUTUALa_entrada: mutual_id_rex (hoja MUTUALES, por mutual_cod_dt);
                           va después de Org. administrador ley 16.744(1152).
      - Porcmutual_entrada : 'Cotización Mutual' de listado_empresas, buscando Empresa_entrada en la
                           columna Nombre (o, si no está, en Empresa); va después de nombreMUTUALa_entrada.
      - Después de Total indemnizaciones no tributables(5565), en orden:
          topeimp_entrada, topeces_entrada, topesalud_entrada, porcAportecaja_entrada,
          aporteAfp_entrada, aporteExpvida_entrada, aporteFappbac_entrada  → Parámetros mensuales por Mes de proceso
          rebajallss_entrada → 3141 + 3143 + 3151 + 3156 + 3158 + 3167 + 3154
    """
    col_rut = _buscar_col_rut(df)
    if col_rut is None:
        raise ErrorArchivo("No se encontró la columna Rut trabajador (1101).")
    valor = df[COL_MES_PROCESO].astype(str) + df[col_rut].astype(str).str.strip()
    df.insert(df.columns.get_loc(col_rut) + 1, COL_MP_RUT, valor)

    # Empresa_entrada / Numcontrato_entrada: búsqueda en listado_empleados por Rut + Fecha inicio contrato
    col_ini = _buscar_col_codigo(df, COL_FECHA_INI_DT)
    if col_ini is None:
        raise ErrorArchivo("No se encontró la columna Fecha inicio contrato (1102).")
    clave_emp = df[col_rut].map(normalizar_rut) + df[col_ini].map(normalizar_fecha)  # ej. 12458987-804/03/2021
    emp_unicos = empleados.drop_duplicates("clave", keep="first").set_index("clave")
    pos = df.columns.get_loc(col_ini)
    df.insert(pos + 1, COL_EMPRESA, clave_emp.map(emp_unicos["Empresa"]).fillna(""))
    num_contrato = clave_emp.map(emp_unicos["Contrato"])
    num_contrato_num = pd.to_numeric(num_contrato, errors="coerce")
    if num_contrato.notna().sum() == num_contrato_num.notna().sum():   # todos numéricos → entero
        num_contrato = num_contrato_num.astype("Int64")
    df.insert(pos + 2, COL_NUM_CONTRATO, num_contrato)
    horas = pd.to_numeric(clave_emp.map(emp_unicos["Horas Semanales"]), errors="coerce")
    df.insert(pos + 3, COL_HORAS_SEM, horas.round(2) if (horas.dropna() % 1 != 0).any() else horas.astype("Int64"))
    df.attrs["empleados_sin_match"] = sorted(clave_emp[df[COL_EMPRESA] == ""].unique().tolist())
    df.attrs["empleados_duplicados"] = sorted(empleados.loc[empleados["clave"].duplicated(), "clave"].unique().tolist())

    col_afp = _buscar_col_codigo(df, COL_AFP_DT)
    if col_afp is None:
        raise ErrorArchivo("No se encontró la columna AFP (1141).")
    tabla_afp = instituciones["AFP"]
    mapa_afp = dict(zip(tabla_afp["afp_cod_dt"], tabla_afp["afp_id_rex"]))
    df.insert(df.columns.get_loc(col_afp) + 1, COL_NOMBRE_AFP, df[col_afp].map(lambda v: mapa_afp.get(_cod(v), "")))

    # Variable de memoria (no se guarda como columna): Mes de proceso + nombreafp_entrada
    clave_afp_hist = (df[COL_MES_PROCESO].astype(str) + df[COL_NOMBRE_AFP].astype(str)).str.lower()
    hist = cot_afp_hist.set_index("id_afp_hist")
    pos = df.columns.get_loc(COL_NOMBRE_AFP)
    # Se expresan en porcentaje (x 100): 0.1145 → 11.45
    df.insert(pos + 1, COL_PORC_AFP, (clave_afp_hist.map(hist["cot_hist_afp"]) * 100).round(4))
    df.insert(pos + 2, COL_SIS, (clave_afp_hist.map(hist["sis_hist"]) * 100).round(4))
    df.attrs["afp_hist_sin_match"] = sorted(
        clave_afp_hist[df[COL_PORC_AFP].isna() & (df[COL_NOMBRE_AFP] != "")].unique().tolist()
    )

    for col_lre, hoja, col_clave, campo, col_nueva in BUSQUEDAS_INSTITUCIONES:
        col_origen = _buscar_col_codigo(df, col_lre)
        if col_origen is None:
            raise ErrorArchivo(f"No se encontró la columna {col_lre}.")
        tabla = instituciones[hoja]
        mapa = dict(zip(tabla[col_clave], tabla[campo]))
        df.insert(df.columns.get_loc(col_origen) + 1, col_nueva,
                  df[col_origen].map(lambda v, mapa=mapa: mapa.get(_cod(v), "")))

    # Porcmutual_entrada: Empresa_entrada → listado_empresas (Nombre o Empresa) → Cotización Mutual
    def _clave(v):
        return "" if pd.isna(v) else str(v).strip().upper()
    mapa_mutual = {}
    for col_busca in ("Empresa", "Nombre"):   # Nombre al final: tiene prioridad si hay coincidencia en ambas
        for k, v in zip(empresas[col_busca], empresas["Cotización Mutual"]):
            if _clave(k):
                mapa_mutual[_clave(k)] = v
    df.insert(df.columns.get_loc(COL_NOMBRE_MUTUAL) + 1, COL_PORC_MUTUAL,
              df[COL_EMPRESA].map(lambda v: mapa_mutual.get(_clave(v))))
    df.attrs["empresas_sin_match"] = sorted(
        {str(v) for v, p in zip(df[COL_EMPRESA], df[COL_PORC_MUTUAL]) if _clave(v) and pd.isna(p)}
    )

    # idEmpresa_entrada: Empresa_entrada → listado_empresas.Nombre → listado_empresas.Empresa; va después de (1102)
    mapa_id = {_clave(n): str(e).strip() for n, e in zip(empresas["Nombre"], empresas["Empresa"])
               if _clave(n) and pd.notna(e)}
    df.insert(df.columns.get_loc(_buscar_col_codigo(df, COL_FECHA_INI_DT)) + 1, COL_ID_EMPRESA,
              df[COL_EMPRESA].map(lambda v: mapa_id.get(_clave(v), "")))
    df.attrs["id_empresa_sin_match"] = sorted(
        {str(v) for v, i in zip(df[COL_EMPRESA], df[COL_ID_EMPRESA]) if _clave(v) and i == ""}
    )

    # Parámetros mensuales por Mes de proceso → después de (5565)
    col_ultima = _buscar_col_codigo(df, COL_ULTIMA_LRE)
    if col_ultima is None:
        raise ErrorArchivo("No se encontró la columna Total indemnizaciones no tributables (5565).")
    par = parametros.set_index("mes_Proc")
    pos = df.columns.get_loc(col_ultima)
    for i, (col_par, col_nueva) in enumerate(PARAMETROS_A_TRAER, start=1):
        df.insert(pos + i, col_nueva, df[COL_MES_PROCESO].map(par[col_par]))

    # rebajallss_entrada: suma de cotizaciones del trabajador que rebajan la base tributable
    cols_rebaja = []
    for cod in CODIGOS_REBAJA_LLSS:
        c = _buscar_col_codigo(df, cod)
        if c is None:
            raise ErrorArchivo(f"No se encontró la columna {cod} para calcular {COL_REBAJA_LLSS}.")
        cols_rebaja.append(c)
    rebaja = df[cols_rebaja].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    df.insert(pos + len(PARAMETROS_A_TRAER) + 1, COL_REBAJA_LLSS, rebaja.round(0).astype("Int64"))

    # sueldoContrato_entrada
    df.insert(df.columns.get_loc(_buscar_col_codigo(df, COL_TASA_IND_DT)) + 1, COL_SUELDO_CONTRATO,
              calcular_sueldo_contrato(df))

    # ultImpSinLic_entrada
    df.insert(df.columns.get_loc(_buscar_col_codigo(df, COL_DIAS_LIC_DT)) + 1, COL_ULT_IMP_SIN_LIC,
              calcular_ult_imp_sin_lic(df))

    # impoMesAntporDdeLic_entrada
    df.insert(df.columns.get_loc(COL_ULT_IMP_SIN_LIC) + 1, COL_IMP_MES_ANT_LIC, calcular_impo_mes_ant_lic(df))

    # imponibleLic_entrada
    df.insert(df.columns.get_loc(COL_IMP_MES_ANT_LIC) + 1, COL_IMPONIBLE_LIC, calcular_imponible_lic(df))
    return df


def calcular_imponible_lic(df):
    """
    imponibleLic_entrada = impoMesAntporDdeLic_entrada
                           + Total haberes imponibles y tributables(5210)
                           + Total haberes imponibles no tributables(5220)   (entero), solo si 1116 > 0.
      - Si 1116 = 0 → 0.
      - Si impoMesAntporDdeLic_entrada está vacío (imponible no encontrado) → vacío (se calcula al ingresarlo a mano).
      - Tope: si el resultado > topeimp_entrada →
              topeimp_entrada - ((ultImpSinLic_entrada / 30) × 1116)   (= topeimp_entrada - impoMesAntporDdeLic_entrada)
        si no, se mantiene el resultado.
        No da negativo: ultImpSinLic_entrada siempre viene topado (mes anterior o sueldo de contrato,
        ambos con min(..., topeimp_entrada)), por lo que impoMesAntporDdeLic_entrada no supera el tope.
    """
    cols = {}
    for cod in (COL_IMP_TRIB_DT, COL_IMP_NO_TRIB_DT):
        cols[cod] = _buscar_col_codigo(df, cod)
        if cols[cod] is None:
            raise ErrorArchivo(f"No se encontró la columna {cod} para calcular {COL_IMPONIBLE_LIC}.")
    haberes = (pd.to_numeric(df[cols[COL_IMP_TRIB_DT]], errors="coerce").fillna(0)
               + pd.to_numeric(df[cols[COL_IMP_NO_TRIB_DT]], errors="coerce").fillna(0))
    impo_lic = pd.to_numeric(df[COL_IMP_MES_ANT_LIC], errors="coerce")
    dias_lic = pd.to_numeric(df[_buscar_col_codigo(df, COL_DIAS_LIC_DT)], errors="coerce").fillna(0)
    tope = pd.to_numeric(df["topeimp_entrada"], errors="coerce")
    resultado = []
    for dl, i, h, t in zip(dias_lic, impo_lic, haberes, tope):
        if dl <= 0:
            resultado.append(0)
        elif pd.isna(i):
            resultado.append(None)
        else:
            valor = i + h
            if pd.notna(t) and valor > t:
                valor = max(t - i, 0)   # i = (ultImpSinLic_entrada / 30) × días licencia(1116); i ≤ tope
            resultado.append(int(round(valor)))
    return pd.Series(resultado, index=df.index, dtype=object)


def calcular_impo_mes_ant_lic(df):
    """
    impoMesAntporDdeLic_entrada = (ultImpSinLic_entrada / 30) × Nro días de licencia médica en el mes(1116)  (entero).
      - Si 1116 = 0 → 0.
      - Si 1116 > 0 y ultImpSinLic_entrada = 'imponible no encontrado' → vacío (se calcula al ingresarlo a mano).
    """
    dias_lic = pd.to_numeric(df[_buscar_col_codigo(df, COL_DIAS_LIC_DT)], errors="coerce").fillna(0)
    ult = pd.to_numeric(df[COL_ULT_IMP_SIN_LIC], errors="coerce")
    resultado = []
    for dl, u in zip(dias_lic, ult):
        if dl <= 0:
            resultado.append(0)
        elif pd.isna(u):
            resultado.append(None)
        else:
            resultado.append(int(round(u / 30 * dl)))
    return pd.Series(resultado, index=df.index, dtype=object)


def calcular_sueldo_contrato(df):
    """
    sueldoContrato_entrada = Sueldo(2101) / Nro días trabajados(1115) × 30  (entero), para todos los registros.
    Si 1115 = 0 (licencia mes completo): se busca hacia atrás, en el mismo Rut, el último mes con 1115 > 0
    y se usa su resultado. Si no hay ninguno → 'sueldo no encontrado'.
    """
    col_rut = _buscar_col_rut(df)
    cols = {}
    for cod in (COL_DIAS_TRAB_DT, COL_SUELDO_DT):
        cols[cod] = _buscar_col_codigo(df, cod)
        if cols[cod] is None:
            raise ErrorArchivo(f"No se encontró la columna {cod} para calcular {COL_SUELDO_CONTRATO}.")
    trab = pd.to_numeric(df[cols[COL_DIAS_TRAB_DT]], errors="coerce").fillna(0)
    sueldo = pd.to_numeric(df[cols[COL_SUELDO_DT]], errors="coerce").fillna(0)
    calculado = (sueldo / trab.where(trab > 0) * 30).round(0)
    rut = df[col_rut].map(normalizar_rut)

    # Por Rut y mes: primer valor calculable
    historia = {}
    for r, mes, v in sorted(zip(rut, df[COL_MES_PROCESO], calculado), key=lambda x: (x[0], x[1])):
        if pd.notna(v):
            historia.setdefault(r, {}).setdefault(mes, v)

    resultado = []
    for r, mes, v in zip(rut, df[COL_MES_PROCESO], calculado):
        if pd.notna(v):
            resultado.append(int(v))
            continue
        valor = TXT_SUELDO_NO_ENCONTRADO
        for mes_ant in sorted((m_ for m_ in historia.get(r, {}) if m_ < mes), reverse=True):
            valor = int(historia[r][mes_ant])
            break
        resultado.append(valor)
    return pd.Series(resultado, index=df.index, dtype=object)


def calcular_ult_imp_sin_lic(df):
    """
    Último mes imponible sin licencia, por registro:
      - Si el registro tiene licencia (1116 > 0): se busca el mismo Rut en el mes anterior; si ese mes
        no tiene licencia (1116 = 0) → min(5210, topeimp_entrada) de ese mes. Si también tiene licencia,
        o el Rut no aparece ese mes, se sigue hacia atrás.
        Si no se encuentra ninguno → min(sueldoContrato_entrada, topeimp_entrada) del mismo registro.
        Si el sueldo tampoco se encontró ('sueldo no encontrado') → 'imponible no encontrado'
        (se resuelve al ingresar el sueldo de contrato a mano en la ventana emergente; también se topa).
      - Si el registro no tiene licencia (1116 = 0): queda vacío.
    """
    col_rut = _buscar_col_rut(df)
    col_lic = _buscar_col_codigo(df, COL_DIAS_LIC_DT)
    col_imp = _buscar_col_codigo(df, COL_IMP_TRIB_DT)
    for c, n in ((col_lic, "1116"), (col_imp, "5210")):
        if c is None:
            raise ErrorArchivo(f"No se encontró la columna ({n}) para calcular {COL_ULT_IMP_SIN_LIC}.")
    rut = df[col_rut].map(normalizar_rut)
    dias_lic = pd.to_numeric(df[col_lic], errors="coerce").fillna(0)
    imponible = pd.to_numeric(df[col_imp], errors="coerce").fillna(0)
    tope = pd.to_numeric(df["topeimp_entrada"], errors="coerce")
    valor_mes = imponible.where(tope.isna(), imponible.clip(upper=tope))   # min(5210, tope)

    # Por Rut y mes: primer registro del mes (días de licencia y valor)
    historia = {}
    for r, mes, dl, v in sorted(zip(rut, df[COL_MES_PROCESO], dias_lic, valor_mes), key=lambda x: (x[0], x[1])):
        historia.setdefault(r, {}).setdefault(mes, (dl, v))

    sueldo = pd.to_numeric(df[COL_SUELDO_CONTRATO], errors="coerce")   # 'sueldo no encontrado' → NaN

    resultado = []
    for r, mes, dl, sc, t in zip(rut, df[COL_MES_PROCESO], dias_lic, sueldo, tope):
        if dl <= 0:
            resultado.append(None)
            continue
        # Sin mes anterior sin licencia → sueldo de contrato del mismo registro, con tope del mes
        if pd.notna(sc):
            valor = int(round(min(sc, t))) if pd.notna(t) else int(sc)
        else:
            valor = TXT_IMP_NO_ENCONTRADO
        for mes_ant in sorted((m_ for m_ in historia.get(r, {}) if m_ < mes), reverse=True):
            dl_ant, v_ant = historia[r][mes_ant]
            if dl_ant == 0:
                valor = int(round(v_ant))
                break
        resultado.append(valor)
    return pd.Series(resultado, index=df.index, dtype=object)


def codigos_sin_match(df, col_codigo, col_resultado):
    """Códigos (únicos) cuya búsqueda quedó vacía, para avisar al usuario."""
    col = _buscar_col_codigo(df, col_codigo)
    if col is None or col_resultado not in df.columns:
        return []
    return sorted({_cod(v) for v, r in zip(df[col], df[col_resultado]) if r == ""})


# ─────────────────────────────────────────────
# 4. EXCEL DE SALIDA
# ─────────────────────────────────────────────
def generar_excel(df, filas_manuales=None):
    """filas_manuales: {columna: [posiciones 0..n-1]} con valores ingresados a mano → celda naranjo claro."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Maestro")
        ws = writer.sheets["Maestro"]
        relleno = PatternFill("solid", fgColor="1A2744")
        fuente = Font(bold=True, color="FFFFFF", size=10)
        relleno_verde = PatternFill("solid", fgColor=VERDE_CLARO)
        fuente_verde = Font(bold=True, color="1A2744", size=10)
        for celda in ws[1]:
            nueva = celda.value in COLUMNAS_NUEVAS
            celda.fill = relleno_verde if nueva else relleno
            celda.font = fuente_verde if nueva else fuente
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[celda.column_letter].width = 16
            if nueva:
                # toda la columna en verde claro para diferenciarla de las originales
                for (c,) in ws.iter_rows(min_row=2, min_col=celda.column, max_col=celda.column):
                    c.fill = relleno_verde
                ws.column_dimensions[celda.column_letter].width = 20
            if celda.value in COLUMNAS_TASA:
                for (c,) in ws.iter_rows(min_row=2, min_col=celda.column, max_col=celda.column):
                    c.number_format = "0.00"
        # Filas con datos no encontrados → fila completa en amarillo
        relleno_amarillo = PatternFill("solid", fgColor=AMARILLO)
        for col_dato, info in DATOS_MANUALES.items():
            if col_dato not in df.columns:
                continue
            for i, v in enumerate(df[col_dato].tolist(), start=2):
                if v == info["texto"]:
                    for c in ws[i]:
                        c.fill = relleno_amarillo
        # Celdas ingresadas a mano → naranjo claro
        relleno_naranjo = PatternFill("solid", fgColor=NARANJO_CLARO)
        for col_dato, posiciones in (filas_manuales or {}).items():
            col_idx = list(df.columns).index(col_dato) + 1
            for pos in posiciones:
                ws.cell(pos + 2, col_idx).fill = relleno_naranjo
        ws.row_dimensions[1].height = 45
        ws.freeze_panes = "D2"   # fija Mes de proceso, Rut trabajador y MpRut_entrada
    return output.getvalue()


# ─────────────────────────────────────────────
# LOG DE ERRORES Y ADVERTENCIAS
# ─────────────────────────────────────────────
class LogProceso:
    """Acumula errores/advertencias (hoja Resumen) y registros afectados (hoja Detalle registros)."""

    def __init__(self):
        self.resumen = []
        self.detalle = []

    def _agregar(self, tipo, etapa, detalle, archivo="", tecnico=""):
        self.resumen.append({
            "Fecha y hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "Tipo": tipo, "Etapa": etapa, "Archivo": archivo,
            "Detalle": detalle, "Detalle técnico": tecnico,
        })

    def error(self, etapa, detalle, archivo="", tecnico=""):
        self._agregar("ERROR", etapa, detalle, archivo, tecnico)

    def advertencia(self, etapa, detalle, archivo=""):
        self._agregar("ADVERTENCIA", etapa, detalle, archivo)

    def registros(self, df, mascara, problema, col_valor=None, etiqueta_valor=""):
        """Agrega al detalle cada fila del consolidado donde 'mascara' es True."""
        col_rut = _buscar_col_rut(df)
        col_ini = _buscar_col_codigo(df, COL_FECHA_INI_DT)
        for _, fila in df[mascara].iterrows():
            self.detalle.append({
                COL_MES_PROCESO: fila.get(COL_MES_PROCESO, ""),
                "Rut trabajador": fila.get(col_rut, "") if col_rut else "",
                "Fecha inicio contrato": fila.get(col_ini, "") if col_ini else "",
                "Problema": problema,
                "Valor buscado": (f"{etiqueta_valor}{fila.get(col_valor, '')}" if col_valor else ""),
            })

    def hay_datos(self):
        return bool(self.resumen or self.detalle)

    def excel(self):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            hojas = {
                "Resumen": pd.DataFrame(self.resumen, columns=["Fecha y hora", "Tipo", "Etapa", "Archivo",
                                                               "Detalle", "Detalle técnico"]),
                "Detalle registros": pd.DataFrame(self.detalle, columns=[COL_MES_PROCESO, "Rut trabajador",
                                                                         "Fecha inicio contrato", "Problema",
                                                                         "Valor buscado"]),
            }
            for nombre, df in hojas.items():
                df.to_excel(writer, index=False, sheet_name=nombre)
                ws = writer.sheets[nombre]
                for celda in ws[1]:
                    celda.fill = PatternFill("solid", fgColor="C53030")
                    celda.font = Font(bold=True, color="FFFFFF", size=10)
                    ws.column_dimensions[celda.column_letter].width = 22
                if nombre == "Resumen":
                    ws.column_dimensions["E"].width = 90
                    ws.column_dimensions["F"].width = 60
                    rojo, amarillo = PatternFill("solid", fgColor="FDE2E2"), PatternFill("solid", fgColor="FEF3C7")
                    for fila in ws.iter_rows(min_row=2):
                        relleno = rojo if fila[1].value == "ERROR" else amarillo
                        for c in fila:
                            c.fill = relleno
                            c.alignment = Alignment(wrap_text=True, vertical="top")
                else:
                    ws.column_dimensions["D"].width = 60
                ws.freeze_panes = "A2"
        return output.getvalue()


def boton_log(log, clave):
    if log.hay_datos():
        n_err = sum(1 for r in log.resumen if r["Tipo"] == "ERROR")
        n_adv = sum(1 for r in log.resumen if r["Tipo"] == "ADVERTENCIA")
        st.download_button(
            f"📄 Descargar log de errores ({n_err} error(es), {n_adv} advertencia(s), "
            f"{len(log.detalle)} registro(s) afectados)",
            data=log.excel(),
            file_name=f"log_errores_lre_dt_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=clave,
        )


def abortar(log, etapa, mensaje, archivo="", excepcion=None):
    """Muestra el error, lo registra en el log, ofrece la descarga del log y detiene el proceso."""
    tecnico = ""
    if excepcion is not None and not isinstance(excepcion, ErrorArchivo):
        tecnico = "".join(traceback.format_exception(type(excepcion), excepcion, excepcion.__traceback__))[-3000:]
    log.error(etapa, mensaje, archivo, tecnico)
    st.error(f"❌ {'Archivo `' + archivo + '`: ' if archivo else ''}{mensaje} Proceso abortado.")
    boton_log(log, "dt_log_abort")
    st.stop()


def advertir(log, etapa, mensaje_pantalla, mensaje_log=None, archivo=""):
    st.warning(f"⚠️ {mensaje_pantalla}")
    log.advertencia(etapa, mensaje_log or mensaje_pantalla.replace("**", ""), archivo)


def _leer_con_nombre(funcion, archivo):
    """Ejecuta un lector y, si falla, deja el nombre del archivo en la excepción."""
    try:
        return funcion(archivo)
    except Exception as e:  # noqa: BLE001
        e.archivo_origen = getattr(archivo, "name", "")
        raise


# ─────────────────────────────────────────────
# INGRESO MANUAL DE IMPONIBLES NO ENCONTRADOS (ventana emergente)
# ─────────────────────────────────────────────
def tabla_pendientes(df):
    """Una fila por (dato, Rut) con registros 'no encontrado' (meses y cantidad de registros)."""
    col_rut = _buscar_col_rut(df)
    filas = []
    for col_dato, info in DATOS_MANUALES.items():
        if not info.get("pedir", True):
            continue
        sin = df[df[col_dato] == info["texto"]]
        for rut, g in sin.groupby(sin[col_rut].map(normalizar_rut), sort=True):
            filas.append({"Dato": col_dato, "Etiqueta": info["etiqueta"], "Rut trabajador": rut,
                          "Meses": ", ".join(sorted(g[COL_MES_PROCESO].unique())), "Registros": len(g)})
    return pd.DataFrame(filas, columns=["Dato", "Etiqueta", "Rut trabajador", "Meses", "Registros"])


def aplicar_manuales(df, manuales):
    """
    manuales: {'columna|rut': valor}. Reemplaza el texto 'no encontrado' de esa columna y Rut por el valor.
    Retorna (df, {columna: [posiciones modificadas]}).
    """
    if not manuales:
        return df, {}
    df = df.copy()
    ruts = df[_buscar_col_rut(df)].map(normalizar_rut)
    modificadas = {}
    for clave, valor in manuales.items():
        col_dato, rut = clave.split("|", 1)
        mascara = (df[col_dato] == DATOS_MANUALES[col_dato]["texto"]) & (ruts == rut)
        if mascara.any():
            df.loc[mascara, col_dato] = int(valor)
            modificadas.setdefault(col_dato, []).extend(i for i, m_ in enumerate(mascara.tolist()) if m_)
    if COL_SUELDO_CONTRATO in modificadas and COL_ULT_IMP_SIN_LIC in df.columns:
        # El sueldo ingresado a mano completa ultImpSinLic_entrada (cuando no hay mes anterior sin licencia)
        antes = (df[COL_ULT_IMP_SIN_LIC] == TXT_IMP_NO_ENCONTRADO).tolist()
        df[COL_ULT_IMP_SIN_LIC] = calcular_ult_imp_sin_lic(df)
        despues = (df[COL_ULT_IMP_SIN_LIC] == TXT_IMP_NO_ENCONTRADO).tolist()
        resueltas = [i for i, (a, d) in enumerate(zip(antes, despues)) if a and not d]
        if resueltas:
            modificadas.setdefault(COL_ULT_IMP_SIN_LIC, []).extend(resueltas)
    if COL_IMP_MES_ANT_LIC in df.columns:
        df[COL_IMP_MES_ANT_LIC] = calcular_impo_mes_ant_lic(df)
    if COL_IMPONIBLE_LIC in df.columns:
        df[COL_IMPONIBLE_LIC] = calcular_imponible_lic(df)
    return df, modificadas


def _dialogo(titulo):
    """st.dialog (ventana emergente) si la versión de Streamlit lo tiene; si no, experimental_dialog."""
    deco = getattr(st, "dialog", None) or getattr(st, "experimental_dialog", None)
    return deco(titulo, width="large") if deco else None


def ventana_manual(pendientes):
    """Ventana emergente para ingresar a mano los datos no encontrados (imponible / sueldo)."""
    guardados = st.session_state.get(CLAVE_IMP_MANUAL, {})

    def contenido():
        st.markdown("Ingresa los datos que el programa no pudo encontrar. "
                    "Cada valor se aplica a todos los registros de ese Rut. Deja en blanco los que no tengas.")
        with st.form("form_manuales", border=False):
            valores = {}
            for etiqueta, grupo in pendientes.groupby("Etiqueta", sort=False):
                st.markdown(f"#### {etiqueta}")
                for _, fila in grupo.iterrows():
                    clave = f"{fila['Dato']}|{fila['Rut trabajador']}"
                    c_rut, c_val = st.columns([3, 2])
                    c_rut.markdown(f"**{fila['Rut trabajador']}**  \n{fila['Registros']} registro(s): {fila['Meses']}")
                    valores[clave] = c_val.number_input(
                        f"{etiqueta} {fila['Rut trabajador']}", min_value=0, step=1, value=guardados.get(clave),
                        placeholder="Ej: 850000", label_visibility="collapsed", key=f"manual_{clave}",
                    )
            c1, c2 = st.columns(2)
            guardar = c1.form_submit_button("💾 Guardar", type="primary", use_container_width=True)
            borrar = c2.form_submit_button("🗑️ Borrar ingresos", use_container_width=True)
        if guardar:
            st.session_state[CLAVE_IMP_MANUAL] = {k: int(v) for k, v in valores.items() if v is not None and v > 0}
            st.rerun()
        if borrar:
            st.session_state[CLAVE_IMP_MANUAL] = {}
            for k in valores:
                st.session_state.pop(f"manual_{k}", None)
            st.rerun()

    deco = _dialogo("✍️ Datos no encontrados")
    if deco:
        deco(contenido)()
    else:   # Streamlit antiguo: se muestra en la página
        with st.container(border=True):
            contenido()


# ─────────────────────────────────────────────
# SELECCIÓN DE MESES A PROCESAR
# ─────────────────────────────────────────────
def selector_rango_meses(meses):
    """
    El usuario elige un rango de meses corridos (desde – hasta) entre los meses subidos.
    Un solo mes = desde y hasta iguales; todos = rango completo (valor por defecto).
    Retorna (mes_desde, mes_hasta).
    """
    if len(meses) == 1:
        st.markdown(f"📅 **Mes a procesar:** {meses[0]}")
        return meses[0], meses[0]
    desde, hasta = st.select_slider(
        "📅 Meses a procesar (desde – hasta)",
        options=meses,
        value=(meses[0], meses[-1]),
        key=f"rango_meses_{meses[0]}_{meses[-1]}_{len(meses)}",
        help="Rango de meses corridos. Para un solo mes, deja ambos extremos en el mismo mes. "
             "Los meses subidos fuera del rango se usan solo como historia (meses anteriores).",
    )
    return desde, hasta


def filtrar_rango(df, en_rango, filas_manuales):
    """
    Deja en el maestro solo los registros del rango elegido. Los cálculos ya se hicieron con todos los
    meses subidos (historia). Reajusta las posiciones de celdas manuales y las listas de advertencias.
    """
    posiciones = [i for i, m_ in enumerate(en_rango.tolist()) if m_]
    nueva_pos = {p: j for j, p in enumerate(posiciones)}
    filas = {c: [nueva_pos[p] for p in ps if p in nueva_pos] for c, ps in (filas_manuales or {}).items()}
    filas = {c: ps for c, ps in filas.items() if ps}

    attrs = dict(df.attrs)
    out = df[en_rango.values].reset_index(drop=True)

    col_rut, col_ini = _buscar_col_rut(out), _buscar_col_codigo(out, COL_FECHA_INI_DT)
    sin_emp = out[COL_EMPRESA] == ""
    attrs["empleados_sin_match"] = sorted(
        (out.loc[sin_emp, col_rut].map(normalizar_rut) + out.loc[sin_emp, col_ini].map(normalizar_fecha)).unique().tolist()
    )
    empresas_rango = {str(v) for v in out[COL_EMPRESA]}
    for clave in ("id_empresa_sin_match", "empresas_sin_match"):
        attrs[clave] = [e for e in attrs.get(clave, []) if e in empresas_rango]
    sin_hist = out[COL_PORC_AFP].isna() & (out[COL_NOMBRE_AFP] != "")
    attrs["afp_hist_sin_match"] = sorted(
        (out.loc[sin_hist, COL_MES_PROCESO].astype(str) + out.loc[sin_hist, COL_NOMBRE_AFP].astype(str))
        .str.lower().unique().tolist()
    )
    out.attrs = attrs
    return out, filas


def nombre_archivo_maestro(mes_desde, mes_hasta, meses_subidos):
    """
    maestro_lre_dt_<periodo>_<fecha hora>.xlsx
      - Todos los meses subidos: todos_2026  (todos los archivos son del mismo año)
      - Un mes:                   2026-03
      - Rango:                    2026-01_a_2026-03
    """
    if mes_desde == meses_subidos[0] and mes_hasta == meses_subidos[-1]:
        periodo = f"todos_{meses_subidos[0][:4]}"
    elif mes_desde == mes_hasta:
        periodo = mes_desde
    else:
        periodo = f"{mes_desde}_a_{mes_hasta}"
    return f"maestro_lre_dt_{periodo}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"


# ─────────────────────────────────────────────
# INTERFAZ
# ─────────────────────────────────────────────
def main():
    st.set_page_config(page_title="Migración desde LRE DT", page_icon="🏛️", layout="wide")
    st.title("🏛️ Migración desde LRE DT")
    st.caption("Etapa 1 — Armado del maestro: consolida los LRE descargados desde la DT y los completa con los datos de referencia.")
    log = LogProceso()

    # Tabla fija de instituciones (data/Instituciones.xlsx)
    try:
        instituciones = cargar_instituciones()
    except Exception as e:  # noqa: BLE001
        abortar(log, "Carga de Instituciones.xlsx", str(e), "Instituciones.xlsx", e)
    with st.expander(
        "🏦 Instituciones cargadas (data/Instituciones.xlsx): "
        + " · ".join(f"{h} {len(d)}" for h, d in instituciones.items())
    ):
        for hoja, df_inst in instituciones.items():
            st.markdown(f"**{hoja}**")
            st.dataframe(df_inst, use_container_width=True, hide_index=True)

    # Archivos de entrada del cliente
    col1, col2 = st.columns(2)
    with col1:
        archivos = st.file_uploader(
            "1️⃣ Archivos CSV descargados desde la DT (todos)",
            type=["csv"],
            accept_multiple_files=True,
            help="El nombre de cada archivo debe incluir año-mes. Ej: Ene_26_declaracion.csv, 2026-01.csv, 202601.csv",
        )
        archivo_empleados = st.file_uploader(
            "2️⃣ Listado de empleados (Rex+)", type=["xlsx"], key="up_empleados",
            help="Se busca por Rut + Fecha inicio contrato para traer Empresa (col. BG) y Contrato (col. BE).",
        )
    with col2:
        archivo_empresas = st.file_uploader("3️⃣ Listado de empresas (Rex+)", type=["xlsx"], key="up_empresas")
        archivo_cot_hist = st.file_uploader(
            "4️⃣ Cotizaciones históricas AFP y SIS", type=["xlsx"], key="up_cot_hist",
            help="Columnas requeridas: id_afp_hist, cot_hist_afp, sis_hist.",
        )
        archivo_parametros = st.file_uploader(
            "5️⃣ Parámetros mensuales", type=["xlsx"], key="up_parametros",
            help="Una fila por mes (mes_Proc = yyyy-mm) con topes AFP, cesantía y salud.",
        )

    faltantes = [n for n, f in [("Archivos CSV de la DT", archivos), ("Listado de empleados", archivo_empleados),
                                ("Listado de empresas", archivo_empresas),
                                ("Cotizaciones históricas AFP y SIS", archivo_cot_hist),
                                ("Parámetros mensuales", archivo_parametros)] if not f]
    if faltantes:
        st.info("Esperando archivos: " + ", ".join(f"**{n}**" for n in faltantes))
        return

    # Archivos de referencia
    try:
        empleados = _leer_con_nombre(cargar_empleados, archivo_empleados)
        empresas = _leer_con_nombre(cargar_empresas, archivo_empresas)
        cot_afp_hist = _leer_con_nombre(cargar_cot_afp_hist, archivo_cot_hist)
        parametros = _leer_con_nombre(cargar_parametros, archivo_parametros)
    except Exception as e:  # noqa: BLE001
        abortar(log, "Lectura de archivos de referencia",
                f"No se pudo leer el archivo: {e}", getattr(e, "archivo_origen", ""), e)
    usadas = empleados.attrs.get("columnas_usadas", {})
    st.caption(
        f"👥 Empleados: {len(empleados)} registros (columnas usadas: "
        + ", ".join(f"{k} → '{v}'" for k, v in usadas.items())
        + f") · 🏢 Empresas: {len(empresas)} · 📈 Cotizaciones AFP/SIS: {len(cot_afp_hist)}"
        + f" · 📅 Parámetros: {len(parametros)} meses"
    )

    # 1) Validar nombres — si uno falla, se aborta todo
    validos, invalidos = validar_nombres(archivos)
    if invalidos:
        for n in invalidos:
            log.error("Validación de nombres", f"{MSG_NOMBRE_INVALIDO}: no incluye año-mes en el nombre.", n)
        abortar(log, "Validación de nombres",
                f"**{MSG_NOMBRE_INVALIDO}.** Archivos sin año-mes en el nombre: " + ", ".join(f"`{n}`" for n in invalidos))

    # 1b) Todos los archivos deben ser del mismo año — si no, se aborta todo
    por_anio = {}
    for f, periodo in validos:
        por_anio.setdefault(periodo[:4], []).append(f.name)
    if len(por_anio) > 1:
        for anio, nombres in sorted(por_anio.items()):
            log.error("Validación de año", f"{MSG_ANIOS_DISTINTOS}: archivo(s) del año {anio}.", ", ".join(nombres))
        abortar(log, "Validación de año",
                f"**{MSG_ANIOS_DISTINTOS}.** Sube solo archivos de un mismo año. "
                + " · ".join(f"{anio}: " + ", ".join(f"`{n}`" for n in sorted(nombres))
                             for anio, nombres in sorted(por_anio.items())))

    # 2) Leer cada CSV (para identificar el archivo que falla) y consolidar
    for f, _ in validos:
        try:
            leer_archivo_dt(f)
        except Exception as e:  # noqa: BLE001
            abortar(log, "Lectura de CSV de la DT", f"Error en la estructura del archivo: {e}", f.name, e)
    try:
        df_cons, resumen = consolidar(validos, instituciones, cot_afp_hist, empleados, empresas, parametros)
    except Exception as e:  # noqa: BLE001
        abortar(log, "Consolidación", f"No se pudieron consolidar los archivos: {e}",
                getattr(e, "archivo_origen", ""), e)

    df_resumen = pd.DataFrame(resumen)
    meses_subidos = sorted(df_cons[COL_MES_PROCESO].unique())

    # Rango de meses a procesar (los meses fuera del rango solo sirven como historia)
    mes_desde, mes_hasta = selector_rango_meses(meses_subidos)
    en_rango = df_cons[COL_MES_PROCESO].between(mes_desde, mes_hasta)
    df_resumen["En el maestro"] = df_resumen[COL_MES_PROCESO].between(mes_desde, mes_hasta).map(
        {True: "Sí", False: "No (solo historia)"})
    faltan_csv = [str(p) for p in pd.period_range(mes_desde, mes_hasta, freq="M")
                  if str(p) not in set(meses_subidos)]

    # Datos no encontrados: se piden solo para el rango; se aplican sobre todos los meses (historia)
    pendientes = tabla_pendientes(df_cons[en_rango])
    claves_pend = {f"{d}|{r}" for d, r in zip(pendientes["Dato"], pendientes["Rut trabajador"])}
    manuales = {k: v for k, v in st.session_state.get(CLAVE_IMP_MANUAL, {}).items() if k in claves_pend}
    df_cons, filas_manuales = aplicar_manuales(df_cons, manuales)
    df_cons, filas_manuales = filtrar_rango(df_cons, en_rango, filas_manuales)
    meses = sorted(df_cons[COL_MES_PROCESO].unique())

    if faltan_csv:
        advertir(log, "Selección de meses",
                 f"El rango {mes_desde} a {mes_hasta} incluye meses sin CSV subido: **{', '.join(faltan_csv)}**.")

    # 3) Advertencias (en pantalla y en el log)
    for nombre_arch, cols in df_cons.attrs.get("columnas_descartadas", []):
        advertir(log, "Lectura de CSV de la DT",
                 f"`{nombre_arch}` trae columnas que no son de la DT y que genera el programa: "
                 f"**{', '.join(cols)}**. Se descartaron las del archivo y se usaron las calculadas.",
                 archivo=nombre_arch)
    duplicados = sorted(df_resumen.loc[df_resumen[COL_MES_PROCESO].duplicated(), COL_MES_PROCESO].unique())
    for mm in duplicados:
        arch = ", ".join(df_resumen.loc[df_resumen[COL_MES_PROCESO] == mm, "Archivo"])
        advertir(log, "Consolidación", f"Hay más de un archivo para **{mm}**: {arch}. Revisa si alguno es una copia.")

    emp_dup = df_cons.attrs.get("empleados_duplicados", [])
    if emp_dup:
        advertir(log, "Cruce con listado de empleados",
                 f"{len(emp_dup)} combinación(es) Rut + Fecha inicio contrato repetidas en el listado de empleados; "
                 f"se usó la primera: **{', '.join(emp_dup[:10])}**{' …' if len(emp_dup) > 10 else ''}",
                 f"{len(emp_dup)} combinación(es) Rut + Fecha inicio repetidas (se usó la primera): {', '.join(emp_dup)}",
                 archivo_empleados.name)
    emp_sin = df_cons.attrs.get("empleados_sin_match", [])
    if emp_sin:
        advertir(log, "Cruce con listado de empleados",
                 f"{len(emp_sin)} combinación(es) Rut + Fecha inicio contrato no están en el listado de empleados — "
                 f"{COL_EMPRESA} y {COL_NUM_CONTRATO} quedaron vacías: **{', '.join(emp_sin[:10])}**"
                 f"{' …' if len(emp_sin) > 10 else ''}",
                 f"{len(emp_sin)} combinación(es) Rut + Fecha inicio no encontradas: {', '.join(emp_sin)}",
                 archivo_empleados.name)
        log.registros(df_cons, df_cons[COL_EMPRESA] == "",
                      f"Rut + Fecha inicio contrato no existe en el listado de empleados ({COL_EMPRESA} y {COL_NUM_CONTRATO} vacías)")

    id_sin = df_cons.attrs.get("id_empresa_sin_match", [])
    if id_sin:
        advertir(log, "Cruce con listado de empresas",
                 f"Empresas que no están en la columna Nombre del listado de empresas: **{', '.join(id_sin)}** — "
                 f"{COL_ID_EMPRESA} quedó vacía para esos registros.", archivo=archivo_empresas.name)
        log.registros(df_cons, df_cons[COL_EMPRESA].isin(id_sin),
                      f"Empresa no existe en Nombre del listado de empresas ({COL_ID_EMPRESA} vacía)",
                      COL_EMPRESA, "Empresa: ")

    emp_sin_mutual = df_cons.attrs.get("empresas_sin_match", [])
    if emp_sin_mutual:
        advertir(log, "Cruce con listado de empresas",
                 f"Empresas que no están en el listado de empresas: **{', '.join(emp_sin_mutual)}** — "
                 f"{COL_PORC_MUTUAL} quedó vacía para esos registros.", archivo=archivo_empresas.name)
        log.registros(df_cons, df_cons[COL_EMPRESA].isin(emp_sin_mutual),
                      f"Empresa no existe en el listado de empresas ({COL_PORC_MUTUAL} vacía)", COL_EMPRESA, "Empresa: ")

    # Datos no encontrados: ingreso manual (ventana emergente) — ya aplicados arriba
    for clave, v in manuales.items():
        col_dato, rut = clave.split("|", 1)
        log.advertencia(DATOS_MANUALES[col_dato]["etiqueta"],
                        f"{col_dato} ingresado manualmente para Rut {rut}: {v:,}".replace(",", "."))
    for col_dato, posiciones in filas_manuales.items():
        mascara_man = pd.Series(False, index=df_cons.index)
        mascara_man.iloc[posiciones] = True
        log.registros(df_cons, mascara_man, f"{col_dato} ingresado manualmente", col_dato, "Valor: ")

    if not pendientes.empty:
        faltan = sum(int((df_cons[c] == i["texto"]).sum()) for c, i in DATOS_MANUALES.items())
        resumen_pend = ", ".join(f"{g['Rut trabajador'].nunique()} Rut sin {e.lower()}"
                                 for e, g in pendientes.groupby("Etiqueta", sort=False))
        col_a, col_b = st.columns([3, 1])
        col_a.info(f"✍️ {resumen_pend} — {len(manuales)} dato(s) ingresado(s) a mano, "
                   f"{faltan} registro(s) aún pendientes.")
        if col_b.button("✍️ Ingresar datos faltantes", use_container_width=True, key="btn_manual"):
            ventana_manual(pendientes)

    for col_dato, info in DATOS_MANUALES.items():
        sin = df_cons[col_dato] == info["texto"]
        if sin.any():
            motivo = ("con licencia sin un mes anterior sin licencia y sin sueldo de contrato"
                      if col_dato == COL_ULT_IMP_SIN_LIC
                      else "con días trabajados = 0 sin un mes anterior con días trabajados")
            advertir(log, info["etiqueta"],
                     f"{int(sin.sum())} registro(s) {motivo} — {col_dato} = '{info['texto']}' "
                     f"(filas marcadas en amarillo en el Excel).")
            log.registros(df_cons, sin, f"{col_dato} = '{info['texto']}'",
                          _buscar_col_codigo(df_cons, COL_DIAS_LIC_DT), "Días licencia: ")

    meses_sin_param = [mm for mm in meses if mm not in set(parametros["mes_Proc"])]
    if meses_sin_param:
        advertir(log, "Parámetros mensuales",
                 f"Meses sin fila en Parámetros mensuales: **{', '.join(meses_sin_param)}** — "
                 f"las columnas de parámetros quedaron vacías para esos meses.",
                 archivo=archivo_parametros.name)
        log.registros(df_cons, df_cons[COL_MES_PROCESO].isin(meses_sin_param),
                      "Mes de proceso no existe en Parámetros mensuales (columnas de parámetros vacías)",
                      COL_MES_PROCESO, "Mes: ")

    hist_sin_match = df_cons.attrs.get("afp_hist_sin_match", [])
    if hist_sin_match:
        advertir(log, "Cruce con cotizaciones históricas AFP/SIS",
                 f"Mes + AFP sin tasa en cotizaciones históricas: **{', '.join(hist_sin_match)}** — "
                 f"{COL_PORC_AFP} y {COL_SIS} quedaron vacías para esos registros.",
                 archivo=archivo_cot_hist.name)
        log.registros(df_cons, df_cons[COL_PORC_AFP].isna() & (df_cons[COL_NOMBRE_AFP] != ""),
                      f"Mes + AFP no existe en cotizaciones históricas ({COL_PORC_AFP} y {COL_SIS} vacías)",
                      COL_NOMBRE_AFP, "AFP: ")

    for col_lre, hoja, _, _, col_nueva in [(COL_AFP_DT, "AFP", None, None, COL_NOMBRE_AFP)] + BUSQUEDAS_INSTITUCIONES:
        sin_match = codigos_sin_match(df_cons, col_lre, col_nueva)
        if sin_match:
            advertir(log, "Cruce con Instituciones.xlsx",
                     f"Códigos {col_lre} que no existen en Instituciones.xlsx (hoja {hoja}): "
                     f"**{', '.join(sin_match)}** — la columna {col_nueva} quedó vacía para esos registros.",
                     archivo="Instituciones.xlsx")
            col_orig = _buscar_col_codigo(df_cons, col_lre)
            log.registros(df_cons, df_cons[col_nueva] == "",
                          f"Código {col_lre} no existe en Instituciones.xlsx hoja {hoja} ({col_nueva} vacía)",
                          col_orig, "Código: ")

    st.success(
        f"✅ **Maestro armado** — {len(df_cons)} registros — meses: **{', '.join(meses)}**"
        + (f" (de {len(meses_subidos)} mes(es) subidos; el resto se usó solo como historia)"
           if len(meses) < len(meses_subidos) else "")
    )
    st.dataframe(df_resumen, use_container_width=True, hide_index=True)

    with st.expander("👁️ Vista previa del maestro"):
        previa = df_cons.head(100).copy()
        for col_dato in list(DATOS_MANUALES) + [COL_IMP_MES_ANT_LIC, COL_IMPONIBLE_LIC]:
            previa[col_dato] = previa[col_dato].map(lambda v: "" if v is None else str(v))
        vista = previa.style.set_properties(
            subset=[c for c in COLUMNAS_NUEVAS if c in df_cons.columns],
            **{"background-color": f"#{VERDE_CLARO}"}
        )
        st.dataframe(vista, use_container_width=True, hide_index=True)

    st.download_button(
        "⬇️ Descargar maestro (.xlsx)",
        data=generar_excel(df_cons, filas_manuales),
        file_name=nombre_archivo_maestro(mes_desde, mes_hasta, meses_subidos),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    boton_log(log, "dt_log_final")


if __name__ == "__main__":
    main()
