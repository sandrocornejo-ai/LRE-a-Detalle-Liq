import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import io
from datetime import datetime

# ─────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────
DATA_DIR = "data"  # Carpeta con archivos de referencia

# ─────────────────────────────────────────────
# ESTILOS REX+
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Rex+ | Liquidaciones en detalle desde LRE",
    page_icon="💼",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Header */
.rex-header {
    background: linear-gradient(90deg, #2d4a7a 0%, #3a5a8f 100%);
    padding: 14px 28px;
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 28px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.2);
}
.rex-logo {
    background: white;
    color: #2d4a7a;
    font-weight: 800;
    font-size: 15px;
    padding: 6px 12px;
    border-radius: 8px;
    letter-spacing: 0.5px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.15);
}
.rex-logo span { color: #00c8e6; }
.rex-divider-header {
    width: 1px;
    height: 28px;
    background: rgba(255,255,255,0.3);
    margin: 0 16px;
}
.rex-title { color: white; font-size: 17px; font-weight: 500; letter-spacing: 0.2px; }
.rex-badge {
    background: #00c8e6;
    color: white;
    font-size: 11px;
    font-weight: 700;
    padding: 5px 14px;
    border-radius: 20px;
    letter-spacing: 1.5px;
}

/* Cards */
.step-card {
    background: white;
    border: 1px solid #e8edf5;
    border-radius: 10px;
    padding: 18px 20px;
    margin-bottom: 12px;
}
.step-label {
    color: #00b4d8;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 4px;
}
.step-title { font-size: 15px; font-weight: 600; color: #1a2744; margin-bottom: 4px; }
.step-desc { font-size: 13px; color: #6b7a9a; }

/* Section title */
.section-title {
    font-size: 20px;
    font-weight: 700;
    color: #1a2744;
    margin-bottom: 4px;
}
.section-sub { font-size: 13px; color: #6b7a9a; margin-bottom: 20px; }

/* Alerts */
.alert-error {
    background: #fff0f0;
    border-left: 4px solid #e53e3e;
    border-radius: 6px;
    padding: 12px 16px;
    margin: 8px 0;
    font-size: 13px;
    color: #c53030;
}
.alert-success {
    background: #f0fff4;
    border-left: 4px solid #38a169;
    border-radius: 6px;
    padding: 12px 16px;
    margin: 8px 0;
    font-size: 13px;
    color: #276749;
}
.alert-warning {
    background: #fffbf0;
    border-left: 4px solid #d69e2e;
    border-radius: 6px;
    padding: 12px 16px;
    margin: 8px 0;
    font-size: 13px;
    color: #744210;
}

/* Divider */
.rex-divider {
    border: none;
    border-top: 1px solid #e8edf5;
    margin: 24px 0;
}

/* Log table */
.log-header {
    background: #1a2744;
    color: white;
    font-size: 13px;
    font-weight: 600;
    padding: 10px 14px;
    border-radius: 8px 8px 0 0;
}

/* Buttons override */
.stButton > button {
    background-color: #1a2744 !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    padding: 8px 20px !important;
    font-size: 14px !important;
}
.stButton > button:hover {
    background-color: #00b4d8 !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTES DE COLUMNAS
# ─────────────────────────────────────────────
COLS_HABERES_AFECTOS = [
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

COLS_HABERES_EXENTOS = [
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

COLS_DESCUENTOS_LEGALES = [
    "Cotización obligatoria previsional (AFP o IPS)(3141)",
    "Cotización obligatoria salud 7%(3143)",
    "Cotización voluntaria para salud(3144)",
    "Cotización AFC - trabajador(3151)",
    "Cotizaciones técnico extranjero para seguridad social fuera de Chile(3146)",
    "Descuento depósito convenido hasta UF 900 anual(3147)",
    "Cotización APVi Mod A(3155)", "Cotización APVi Mod B hasta UF50(3156)",
    "Cotización APVc Mod A(3157)", "Cotización APVc Mod B hasta UF50(3158)",
    "Retención préstamo clase media 2020 (Ley 21.252) (3166)",
    "Rebaja zona extrema DL 889 (3167)",
    "Cotización adicional trabajo pesado - trabajador(3154)"
]

COLS_OTROS_DESCUENTOS = [
    "Cuota sindical 1(3171)", "Cuota sindical 2(3172)", "Cuota sindical 3(3173)",
    "Cuota sindical 4(3174)", "Cuota sindical 5(3175)", "Cuota sindical 6(3176)",
    "Cuota sindical 7(3177)", "Cuota sindical 8(3178)", "Cuota sindical 9(3179)",
    "Cuota sindical 10(3180)", "Crédito social CCAF(3110)",
    "Cuota vivienda o educación(3181)", "Crédito cooperativas de ahorro(3182)",
    "Otros descuentos autorizados y solicitados por el trabajador(3183)",
    "Mayor retención de impuestos solicitada por el trabajador(3163)",
    "Donaciones culturales y de reconstrucción(3184)", "Otros descuentos(3185)",
    "Pensiones de alimentos(3186)", "Descuento mujer casada(3187)",
    "Descuentos por anticipos y préstamos(3188)"
]

COLS_APORTES_EMPLEADOR = [
    "AFC - Aporte empleador(4151)",
    "Aporte empleador seguro accidentes del trabajo y Ley SANNA(4152)",
    "Aporte empleador indemnización a todo evento(4131)",
    "Aporte adicional trabajo pesado - empleador(4154)",
    "Aporte empleador seguro invalidez y sobrevivencia(4155)",
    "APVC - Aporte Empleador(4157)"
]

GRUPOS_AFP = {
    "afp", "reliquidaAfp", "afpAhor", "cesEmpleado", "reliquidaCesEmpl",
    "trabajoPesaEmpl", "voluntarioCoti", "voluntarioAhor", "reliquidaTrabEmpl",
    "trabajoPesa", "reliquidaTrabPesa", "sis", "reliquidaSis", "cesAporteCi",
    "reliquidaCesCi", "cesAporteSol", "reliquidaCesSol", "aporteAFPemp"
}
GRUPOS_ISAPRE = {"isapre", "reliquidaIsapre"}
GRUPOS_MUTUAL = {"mutual", "reliquidaMutual"}
GRUPOS_CCAF = {
    "ccafReliquida", "cajaCred", "cajaDent", "cajaLeas", "cajaVida",
    "cajaOtro", "cajaAhor", "cajaSegu", "cajaComp", "reliquidaCcaf"
}
GRUPOS_AFP_MUTUAL_AFECTO = {
    "afp", "isapre", "reliquidaIsapre", "reliquidaAfp", "afpAhor",
    "trabajoPesaEmpl", "voluntarioCoti", "voluntarioAhor", "reliquidaTrabEmpl",
    "trabajoPesa", "reliquidaTrabPesa", "sis", "reliquidaSis", "aporteAFPemp",
    "mutual", "reliquidaMutual"
}
GRUPOS_CES_AFECTO = {
    "cesEmpleado", "reliquidaCesEmpl", "cesAporteCi",
    "reliquidaCesCi", "cesAporteSol", "reliquidaCesSol"
}

# ─────────────────────────────────────────────
# FUNCIONES DE CARGA DE REFERENCIAS
# ─────────────────────────────────────────────
@st.cache_data
def cargar_referencias():
    refs = {}
    # Archivos con header en fila 1 (tienen título en fila 0)
    header1 = {"listado_empleados", "listado_empresas"}
    archivos = {
        "equiv_conceptos":  "equiv_conceptos.xlsx",
        "listado_empleados": "listado_empleados.xlsx",
        "listado_empresas":  "listado_empresas.xlsx",
        "inst_mutuales":    "inst_mutuales.xlsx",
        "inst_cajas":       "inst_cajas.xlsx",
        "inst_afp":         "inst_afp.xlsx",
        "inst_salud":       "inst_salud.xlsx",
        "cot_afp_hist":     "cot_afp_hist.xlsx",
        "parametros":       "parametrosMesuales.xlsx",
    }
    errores = []
    for key, fname in archivos.items():
        path = os.path.join(DATA_DIR, fname)
        if os.path.exists(path):
            h = 1 if key in header1 else 0
            refs[key] = pd.read_excel(path, header=h)
        else:
            errores.append(fname)
    return refs, errores


# ─────────────────────────────────────────────
# FUNCIONES DE PROCESAMIENTO
# ─────────────────────────────────────────────
def extraer_fecha_proceso(nombre_archivo):
    """Extrae los últimos 6 chars antes de .csv → formato aaaa-mm"""
    base = os.path.splitext(nombre_archivo)[0]
    sufijo = base[-6:]
    return f"{sufijo[:4]}-{sufijo[4:]}"

def safe_sum(df, cols):
    """Suma columnas que existen en el df, ignorando las que no existen."""
    cols_presentes = [c for c in cols if c in df.columns]
    if not cols_presentes:
        return pd.Series(0, index=df.index)
    return df[cols_presentes].fillna(0).sum(axis=1)

def get_col(df, col, default=0):
    """Obtiene una columna del df o retorna default si no existe."""
    if col in df.columns:
        return df[col].fillna(default)
    return pd.Series(default, index=df.index)

def validar_archivos(archivos_subidos):
    """Valida que todos los archivos sean de la misma empresa."""
    nombres = [f.name for f in archivos_subidos]
    prefijos = [n[:10] for n in nombres]
    if len(set(prefijos)) > 1:
        return False, prefijos
    return True, prefijos

def calcular_totales(df):
    """Calcula los 5 totales pre-validación."""
    df = df.copy()
    df["_total_haberes_afectos"] = safe_sum(df, COLS_HABERES_AFECTOS)
    df["_total_haberes_exentos"] = safe_sum(df, COLS_HABERES_EXENTOS)
    df["_total_descuentos_legales"] = safe_sum(df, COLS_DESCUENTOS_LEGALES)
    df["_total_otros_descuentos"] = safe_sum(df, COLS_OTROS_DESCUENTOS)
    df["_total_aportes_empleador"] = safe_sum(df, COLS_APORTES_EMPLEADOR)
    return df

VALIDACIONES_META = {
    "V1": {
        "nombre": "Haberes afectos",
        "descripcion": "El total de haberes imponibles y tributables calculado no coincide con la columna de control del CSV.",
        "formula": "total_haberes_afectos = suma de cols (2101..2199, 2700..2799)",
        "col_ctrl_label": "Total haberes imponibles y tributables (5210)",
    },
    "V2": {
        "nombre": "Haberes exentos",
        "descripcion": "El total de haberes no imponibles y no tributables calculado no coincide con la columna de control del CSV.",
        "formula": "total_haberes_exentos = suma de cols (3100..3199, 3700..3799)",
        "col_ctrl_label": "Total haberes no imponibles y no tributables (5230)",
    },
    "V3": {
        "nombre": "Descuentos legales",
        "descripcion": "El total de descuentos legales calculado no coincide con la columna de control del CSV.",
        "formula": "total_descuentos_legales = suma(3141, 3143, 3144, 3146, 3147, 3151, 3154, 3155, 3156, 3157, 3158, 3166, 3167)",
        "col_ctrl_label": "Total descuentos por cotizaciones del trabajador (5341)",
    },
    "V4": {
        "nombre": "Otros descuentos",
        "descripcion": "El total de otros descuentos calculado no coincide con la columna de control del CSV.",
        "formula": "total_otros_descuentos = suma de cols (5300..5302)",
        "col_ctrl_label": "Total otros descuentos (5302)",
    },
    "V5": {
        "nombre": "Aportes empleador",
        "descripcion": "El total de aportes del empleador calculado no coincide con la columna de control del CSV.",
        "formula": "total_aportes_empleador = suma de cols (5400..5410)",
        "col_ctrl_label": "Total aportes empleador (5410)",
    },
    "V6": {
        "nombre": "Total liquido",
        "descripcion": "El total liquido calculado (haberes - descuentos) no coincide con la columna de control del CSV.",
        "formula": "(haberes_afectos + haberes_exentos) - (desc_legales + otros_desc + Impuesto retenido por remuneraciones(3161)) = Total liquido(5501)",
        "col_ctrl_label": "Total liquido (5501)",
    },
}

def validar_cuadraturas(df, nombre_archivo):
    """Ejecuta las 6 validaciones y retorna lista de errores."""
    errores = []
    tol = 1  # tolerancia de 1 peso por redondeo

    validaciones = [
        ("V1", "_total_haberes_afectos", "Total haberes imponibles y tributables(5210)"),
        ("V2", "_total_haberes_exentos", "Total haberes no imponibles y no tributables(5230)"),
        ("V3", "_total_descuentos_legales", "Total descuentos por cotizaciones del trabajador(5341)"),
        ("V4", "_total_otros_descuentos", "Total otros descuentos(5302)"),
        ("V5", "_total_aportes_empleador", "Total aportes empleador(5410)"),
    ]

    for codigo, col_calc, col_ctrl in validaciones:
        if col_ctrl not in df.columns:
            continue
        meta = VALIDACIONES_META[codigo]
        ctrl = df[col_ctrl].fillna(0)
        calc = df[col_calc].fillna(0)
        mask = (calc - ctrl).abs() > tol
        filas_error = df[mask]
        for _, row in filas_error.iterrows():
            errores.append({
                "Archivo": nombre_archivo,
                "RUT": row.get("Rut trabajador(1101)", "N/D"),
                "Validacion": codigo,
                "Nombre validacion": meta["nombre"],
                "Descripcion": meta["descripcion"],
                "Formula aplicada": meta["formula"],
                "Col control CSV": meta["col_ctrl_label"],
                "Valor Calculado": round(calc[row.name], 2),
                "Valor Control": round(ctrl[row.name], 2),
                "Diferencia": round(calc[row.name] - ctrl[row.name], 2)
            })

    # V6: liquidez
    if "Total líquido(5501)" in df.columns:
        meta = VALIDACIONES_META["V6"]
        liq_calc = (df["_total_haberes_afectos"] + df["_total_haberes_exentos"]) - \
                   (df["_total_descuentos_legales"] + df["_total_otros_descuentos"] + get_col(df, "Impuesto retenido por remuneraciones(3161)"))
        liq_ctrl = df["Total líquido(5501)"].fillna(0)
        mask = (liq_calc - liq_ctrl).abs() > tol
        filas_error = df[mask]
        for _, row in filas_error.iterrows():
            errores.append({
                "Archivo": nombre_archivo,
                "RUT": row.get("Rut trabajador(1101)", "N/D"),
                "Validacion": "V6",
                "Nombre validacion": meta["nombre"],
                "Descripcion": meta["descripcion"],
                "Formula aplicada": meta["formula"],
                "Col control CSV": meta["col_ctrl_label"],
                "Valor Calculado": round(liq_calc[row.name], 2),
                "Valor Control": round(liq_ctrl[row.name], 2),
                "Diferencia": round(liq_calc[row.name] - liq_ctrl[row.name], 2)
            })

    return errores

def generar_filas_salida(df, fecha_proceso, refs):
    """Genera las filas del archivo de salida via pivot de conceptos."""
    filas = []

    # Cargar referencias
    equiv      = refs.get("equiv_conceptos", pd.DataFrame())
    empleados  = refs.get("listado_empleados", pd.DataFrame())
    empresas   = refs.get("listado_empresas", pd.DataFrame())
    mutuales   = refs.get("inst_mutuales", pd.DataFrame())
    cajas      = refs.get("inst_cajas", pd.DataFrame())
    inst_afp   = refs.get("inst_afp", pd.DataFrame())
    inst_salud = refs.get("inst_salud", pd.DataFrame())
    cot_afp    = refs.get("cot_afp_hist", pd.DataFrame())
    params     = refs.get("parametros", pd.DataFrame())

    # Diccionario equiv conceptos
    equiv_dict = {}
    if not equiv.empty and "cod_lre" in equiv.columns and "concepto_detalle" in equiv.columns:
        equiv_dict = dict(zip(equiv["cod_lre"], equiv["concepto_detalle"]))

    # Diccionarios de instituciones (cod_lre numérico → id texto)
    afp_dict    = dict(zip(inst_afp["cod_lre"], inst_afp["id_afp"])) if not inst_afp.empty and "cod_lre" in inst_afp.columns else {}
    salud_dict  = dict(zip(inst_salud["cod_lre"], inst_salud["id_inst"])) if not inst_salud.empty and "cod_lre" in inst_salud.columns else {}
    mutual_dict = dict(zip(mutuales["cod_lre"], mutuales["id_mutual"])) if not mutuales.empty and "cod_lre" in mutuales.columns else {}
    nombre_mutual_dict = dict(zip(mutuales["cod_lre"], mutuales["nombre_mutual"])) if not mutuales.empty else {}
    caja_dict   = dict(zip(cajas["cod_lre"], cajas["id_ccaf"])) if not cajas.empty and "cod_lre" in cajas.columns else {}

    # Diccionario empleados: Rut → {AFP, Isapre, Empresa}
    emp_dict = {}
    if not empleados.empty and "Rut" in empleados.columns:
        for _, er in empleados.iterrows():
            emp_dict[str(er["Rut"])] = {
                "afp": str(er.get("AFP", "") or "").lower().strip(),
                "isapre": str(er.get("Isapre", "") or "").lower().strip(),
                "empresa_nombre": str(er.get("Empresa", "") or "").strip(),
            }

    # Diccionario contratos: Rut → {num_contrato, multiple}
    contrato_dict = {}  # Rut -> numero de contrato
    ruts_multiples = []  # Ruts con más de un contrato
    if not empleados.empty and "Rut" in empleados.columns and "Contrato" in empleados.columns:
        conteos = empleados.groupby("Rut").size()
        ruts_multiples = list(conteos[conteos > 1].index.astype(str))
        unico = empleados[~empleados["Rut"].astype(str).isin(ruts_multiples)]
        contrato_dict = dict(zip(unico["Rut"].astype(str), unico["Contrato"]))

    # Diccionario empresas: Nombre largo → código corto
    empresa_dict = {}
    mutual_cot_dict = {}
    if not empresas.empty and "Nombre" in empresas.columns and "Empresa" in empresas.columns:
        empresa_dict = dict(zip(empresas["Nombre"].astype(str).str.strip(), empresas["Empresa"].astype(str).str.strip()))
    if not empresas.empty and "Mutual" in empresas.columns and "Cotización Mutual" in empresas.columns:
        mutual_cot_dict = dict(zip(empresas["Mutual"].astype(str).str.strip(), empresas["Cotización Mutual"]))

    # Parámetros mensuales
    tope_afp = 0
    tope_ces = 0
    tope_salud = 0
    if not params.empty and "mes_Proc" in params.columns:
        fila_param = params[params["mes_Proc"].astype(str).str[:7] == fecha_proceso]
        if fila_param.empty:
            fila_param = params  # fallback a primera fila
        if "topeImp_pesos_afp" in params.columns:
            tope_afp = float(fila_param["topeImp_pesos_afp"].iloc[0] or 0)
        if "topeCes_pesos" in params.columns:
            tope_ces = float(fila_param["topeCes_pesos"].iloc[0] or 0)
        if "topeSalud_pesos" in params.columns:
            tope_salud = float(fila_param["topeSalud_pesos"].iloc[0] or 0)

    # Columnas de conceptos
    cols_concepto = [c for c in df.columns if c in equiv_dict]

    for _, row in df.iterrows():
        rut             = str(row.get("Rut trabajador(1101)", "") or "").strip()
        cod_afp         = row.get("AFP(1141)", 0) or 0
        cod_salud       = row.get("FONASA - ISAPRE(1143)", 0) or 0
        cod_ccaf        = row.get("CCAF(1110)", 0) or 0
        cod_mutual      = row.get("Org. administrador ley 16.744(1152)", 0) or 0
        col_3110        = row.get("Crédito social CCAF(3110)", 0) or 0
        dias_trabajados = row.get("Nro días trabajados en el mes(1115)", 0) or 0
        dias_licencia   = row.get("Nro días de licencia médica en el mes(1116)", 0) or 0
        dias_vacaciones = row.get("Nro días de vacaciones en el mes(1117)", 0) or 0
        sueldo          = row.get("Sueldo(2101)", 0) or 0
        total_imponible = row.get("Total haberes imponibles y tributables(5210)", 0) or 0
        rebaja_zona     = row.get("Rebaja zona extrema DL 889 (3167)", 0) or 0

        # Resolver instituciones desde códigos numéricos del CSV
        id_afp_emp    = afp_dict.get(int(cod_afp), "") if cod_afp else ""
        id_salud_emp  = salud_dict.get(int(cod_salud), "") if cod_salud else ""
        id_mutual_emp = mutual_dict.get(int(cod_mutual), "") if cod_mutual else ""
        id_ccaf_emp   = caja_dict.get(int(col_3110), "") if col_3110 else ""

        # Empresa desde listado_empleados → listado_empresas
        empresa_salida = ""
        emp_info = emp_dict.get(rut, {})
        nombre_empresa = emp_info.get("empresa_nombre", "")
        if nombre_empresa:
            empresa_salida = empresa_dict.get(nombre_empresa, "")

        # Cotización mutual para esta empresa
        nombre_mutual_emp = nombre_mutual_dict.get(int(cod_mutual), "") if cod_mutual else ""

        monto_init = (sueldo / dias_trabajados * 30) if dias_trabajados > 0 else 0

        # Calcular Total rebajas por LLSS para concepto impuesto
        cot_salud_total = (row.get("Cotización obligatoria salud 7%(3143)", 0) or 0) +                           (row.get("Cotización voluntaria para salud(3144)", 0) or 0)
        salud_rebaja = min(tope_salud, cot_salud_total) if tope_salud > 0 else cot_salud_total
        rebajas_llss_impuesto = (
            (row.get("Cotización obligatoria previsional (AFP o IPS)(3141)", 0) or 0) +
            (row.get("Cotización AFC - trabajador(3151)", 0) or 0) +
            (row.get("Cotización APVi Mod B hasta UF50(3156)", 0) or 0) +
            (row.get("Cotización adicional trabajo pesado - trabajador(3154)", 0) or 0) +
            salud_rebaja
        )

        def make_fila(id_concepto, monto, id_institucion, afecto, cot_jubilacion, rebaja_zona_override=None):
            return {
                "Fecha de proceso":           fecha_proceso,
                "Id empleado":                rut,
                "Número de contrato":         contrato_dict.get(rut, 1),
                "Id del concepto":            id_concepto,
                "Monto del concepto":         monto,
                "Afecto":                     afecto,
                "Id de institución":          id_institucion,
                "Cotización de jubilación":   cot_jubilacion,
                "Días de licencias":          dias_licencia,
                "Días trabajados":            dias_trabajados,
                "Fecha de aplicación":        "x",
                "Empresa":                    empresa_salida,
                "Total de rebajas por LLSS":  rebajas_llss_impuesto if id_concepto == "impuesto" else 0,
                "Rentas no gravadas":         (row.get("_total_haberes_exentos", 0) or 0) if id_concepto == "impuesto" else 0,
                "Rebaja por zona extrema":    rebaja_zona if rebaja_zona_override is None else rebaja_zona_override,
                "Jornada":                    "C",
                "Días de vacaciones":         dias_vacaciones,
                "Monto Init":                 round(monto_init, 2),
                "Fase":                       1,
            }

        # Monto especial isapre = cotizacion obligatoria + voluntaria salud
        monto_isapre = (row.get("Cotización obligatoria salud 7%(3143)", 0) or 0) +                        (row.get("Cotización voluntaria para salud(3144)", 0) or 0)

        # Fila por cada concepto
        conceptos_ya_generados = set()
        for col_csv in cols_concepto:
            monto = row.get(col_csv, 0) or 0
            id_concepto = equiv_dict.get(col_csv, "")
            if monto == 0 and id_concepto not in ("cesEmpleado", "impuesto"):
                continue
            if id_concepto == "isapre":
                if "isapre" in conceptos_ya_generados:
                    continue
                monto = monto_isapre
                if monto == 0:
                    continue
            if id_concepto in conceptos_ya_generados and id_concepto not in ("CuotaSindMi", "otrosHaberes", "otrosDesctoMi"):
                continue

            # Id de institución
            id_institucion = ""
            if id_concepto in GRUPOS_AFP:
                id_institucion = id_afp_emp
            elif id_concepto in GRUPOS_ISAPRE:
                id_institucion = id_salud_emp
            elif id_concepto in GRUPOS_MUTUAL:
                id_institucion = id_mutual_emp
            elif id_concepto in GRUPOS_CCAF and col_3110 != 0:
                id_institucion = id_ccaf_emp

            # Afecto
            afecto = 0
            if id_concepto == "impuesto":
                afecto = (row.get("_total_haberes_afectos", 0) or 0) - rebajas_llss_impuesto
            elif id_concepto == "totalesEmpl":
                afecto = (row.get("_total_haberes_afectos", 0) or 0) + (row.get("_total_haberes_exentos", 0) or 0)
            elif id_concepto in GRUPOS_AFP_MUTUAL_AFECTO:
                afecto = min(total_imponible, tope_afp) if tope_afp > 0 else total_imponible
            elif id_concepto in GRUPOS_CES_AFECTO:
                afecto = min(total_imponible, tope_ces) if tope_ces > 0 else total_imponible

            # Cotización de jubilación
            cot_jubilacion = 0
            if id_concepto == "afp":
                key_afp = f"{fecha_proceso}{id_afp_emp}"
                if not cot_afp.empty and "id_afp_hist" in cot_afp.columns:
                    r = cot_afp[cot_afp["id_afp_hist"] == key_afp]
                    if not r.empty:
                        cot_jubilacion = r.iloc[0].get("cot_hist_afp", 0)
            elif id_concepto == "sis":
                key_sis = f"{fecha_proceso}{id_afp_emp}"
                if not cot_afp.empty and "id_afp_hist" in cot_afp.columns:
                    r = cot_afp[cot_afp["id_afp_hist"] == key_sis]
                    if not r.empty:
                        cot_jubilacion = r.iloc[0].get("sis_hist", 0)
            elif id_concepto == "cesEmpleado":
                cot_jubilacion = 0.6
            elif id_concepto == "isapre":
                cot_jubilacion = monto
            elif id_concepto == "mutual":
                cot_jubilacion = mutual_cot_dict.get(nombre_mutual_emp, 0)
            elif id_concepto == "licenciaDias":
                cot_jubilacion = dias_licencia

            cot_fmt = round(cot_jubilacion * 100, 4) if isinstance(cot_jubilacion, float) and id_concepto not in ("isapre", "licenciaDias", "cesEmpleado") else cot_jubilacion
            conceptos_ya_generados.add(id_concepto)
            filas.append(make_fila(id_concepto, monto, id_institucion, afecto, cot_fmt))

        # Fila adicional licenciaDias si aplica
        if dias_licencia > 0:
            filas.append(make_fila("licenciaDias", dias_licencia, "", 0, 0, rebaja_zona_override=0))

    df_result = pd.DataFrame(filas)
    # Filtrar filas de RUTs con múltiples contratos
    if ruts_multiples and not df_result.empty:
        df_result = df_result[~df_result["Id empleado"].astype(str).isin(ruts_multiples)]
    return df_result, ruts_multiples

def generar_excel(df_salida):
    """Genera el Excel de salida con formato."""
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Migración"

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
# INTERFAZ PRINCIPAL
# ─────────────────────────────────────────────

# Header con botones Limpiar y Salir
st.markdown("""
<style>
div[data-testid="stHorizontalBlock"] > div:nth-child(2) button,
div[data-testid="stHorizontalBlock"] > div:nth-child(3) button {
    margin-top: -58px !important;
    height: 32px !important;
    padding: 0 14px !important;
    font-size: 12px !important;
    border-radius: 6px !important;
}
</style>
<div class="rex-header">
    <div style="display:flex; align-items:center;">
        <div class="rex-logo">Rex<span>+</span></div>
        <div class="rex-divider-header"></div>
        <span class="rex-title">Liquidaciones en detalle desde LRE</span>
    </div>
    <div style="display:flex; align-items:center; gap:12px;">
        <div class="rex-badge">PRODUCCIÓN</div>
    </div>
</div>
""", unsafe_allow_html=True)

_h_spacer, _h_limpiar, _h_salir = st.columns([9.2, 0.9, 0.9])
with _h_limpiar:
    if st.button("🧹 Limpiar", key="btn_limpiar", use_container_width=True):
        for k in ["validacion_ejecutada", "validacion_ok", "errores_validacion",
                  "dfs", "nombre_empresa", "excel_listo", "excel_bytes", "ruts_multiples"]:
            if k in st.session_state:
                del st.session_state[k]
        st.rerun()
with _h_salir:
    if st.button("🚪 Salir", key="btn_salir_header", use_container_width=True):
        st.markdown("""
        <div style="background:#fdecea; border:1px solid #e05252; border-radius:8px;
             padding:12px; font-size:13px; color:#a02020; margin-top:10px;">
            <b>Sesión finalizada.</b> Puedes cerrar esta ventana.
        </div>
        """, unsafe_allow_html=True)
        st.stop()

# Título sección
st.markdown('<div class="section-title">📂 Liquidaciones en detalle desde LRE</div>', unsafe_allow_html=True)
st.markdown('<div class="section-sub">Sube uno o más archivos CSV del mismo RUT empresa para generar el archivo de salida en Excel.</div>', unsafe_allow_html=True)

# Cómo funciona
col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("""
    <div class="step-card">
        <div class="step-label">PASO 1</div>
        <div class="step-title">Subir archivos CSV</div>
        <div class="step-desc">Uno o más archivos del mismo RUT empresa. Ej: 76247825-0_202601.csv</div>
    </div>""", unsafe_allow_html=True)
with col2:
    st.markdown("""
    <div class="step-card">
        <div class="step-label">PASO 2</div>
        <div class="step-title">Validación automática</div>
        <div class="step-desc">Se verifican las cuadraturas contables de cada registro.</div>
    </div>""", unsafe_allow_html=True)
with col3:
    st.markdown("""
    <div class="step-card">
        <div class="step-label">PASO 3</div>
        <div class="step-title">Descargar Excel</div>
        <div class="step-desc">Si todo cuadra, se genera el archivo de salida listo para importar.</div>
    </div>""", unsafe_allow_html=True)

st.markdown('<hr class="rex-divider">', unsafe_allow_html=True)

# Cargar referencias
refs, errores_refs = cargar_referencias()
if errores_refs:
    st.markdown(f'<div class="alert-warning">⚠️ Archivos de referencia no encontrados en <b>/data</b>: {", ".join(errores_refs)}</div>', unsafe_allow_html=True)

# Upload
st.markdown("### 📤 Subir archivos CSV")
archivos = st.file_uploader(
    "Selecciona uno o más archivos CSV de Previred",
    type=["csv"],
    accept_multiple_files=True,
    help="Los archivos deben corresponder al mismo RUT empresa (primeros 10 caracteres del nombre)"
)

if archivos:
    st.markdown(f'<div class="alert-success">✅ {len(archivos)} archivo(s) cargado(s): {", ".join([f.name for f in archivos])}</div>', unsafe_allow_html=True)

    # Validar misma empresa
    valido, prefijos = validar_archivos(archivos)
    if not valido:
        st.markdown(f"""
        <div class="alert-error">
            ❌ <b>Los archivos no corresponden a la misma empresa.</b><br>
            Se detectaron distintos RUT empresa: {", ".join(set(prefijos))}<br>
            Por favor sube solo archivos del mismo RUT empresa.
        </div>""", unsafe_allow_html=True)
        st.stop()

    # Inicializar session_state
    for key, val in [("validacion_ejecutada", False), ("validacion_ok", False),
                     ("errores_validacion", []), ("dfs", []),
                     ("nombre_empresa", ""), ("excel_listo", False), ("excel_bytes", None),
                     ("ruts_multiples", [])]:
        if key not in st.session_state:
            st.session_state[key] = val

    if st.button("▶ Ejecutar validaciones"):
        todos_errores = []
        dfs = []
        with st.spinner("Procesando archivos..."):
            for archivo in archivos:
                df = pd.read_csv(archivo, encoding="latin-1", sep=None, engine="python")
                df = calcular_totales(df)
                errores = validar_cuadraturas(df, archivo.name)
                todos_errores.extend(errores)
                fecha_proceso = extraer_fecha_proceso(archivo.name)
                df["_fecha_proceso"] = fecha_proceso
                dfs.append(df)
        st.session_state.dfs = dfs
        st.session_state.errores_validacion = todos_errores
        st.session_state.validacion_ok = (len(todos_errores) == 0)
        st.session_state.validacion_ejecutada = True
        st.session_state.nombre_empresa = archivos[0].name[:10]
        st.session_state.excel_listo = False
        st.session_state.excel_bytes = None

    # Mostrar resultados siempre que se haya ejecutado
    if st.session_state.get("validacion_ejecutada"):
        st.markdown('<hr class="rex-divider">', unsafe_allow_html=True)
        st.markdown("### 🔍 Resultado de validaciones")

        if st.session_state.errores_validacion:
            errores_list = st.session_state.errores_validacion
            n_errores = len(errores_list)
            # Agrupar por validacion para mostrar cuántos RUT fallan en cada una
            df_err = pd.DataFrame(errores_list)
            grupos = df_err.groupby(["Validacion", "Nombre validacion"])

            st.markdown(f"""
            <div class="alert-error">
                ❌ <b>No se puede generar el archivo de salida.</b><br>
                Se encontraron <b>{n_errores} registro(s) con error</b> en {grupos.ngroups} validacion(es).
            </div>""", unsafe_allow_html=True)

            with st.expander("📋 Ver log de errores detallado", expanded=True):
                for (codigo, nombre), grupo in grupos:
                    ruts = grupo["RUT"].unique()
                    diff_max = grupo["Diferencia"].abs().max()
                    diff_fmt = f"${diff_max:,.0f}".replace(",", ".")

                    st.markdown(f"""
                    <div style="border:1px solid #e05252; border-radius:10px; margin-bottom:16px; overflow:hidden;">
                      <div style="background:#fdf0f0; padding:10px 16px; display:flex; justify-content:space-between; align-items:center;">
                        <span style="font-weight:600; color:#a02020; font-size:14px;">&#10060; {codigo} — {nombre}</span>
                        <span style="font-size:12px; color:#a02020;">{len(grupo)} registro(s) · mayor diferencia: {diff_fmt}</span>
                      </div>
                    </div>""", unsafe_allow_html=True)

                    # Descripcion y formula
                    meta = VALIDACIONES_META.get(codigo, {})
                    st.markdown(f"""
                    <div style="margin:-12px 0 4px 4px; padding:10px 16px; border-left:3px solid #e05252; background:transparent;">
                      <p style="font-size:13px; color:#555; margin:0 0 6px;">{meta.get('descripcion','')}</p>
                      <code style="font-size:12px;">{meta.get('formula','')}</code>
                    </div>""", unsafe_allow_html=True)

                    # Tabla de registros con error
                    tabla = grupo[["RUT", "Archivo", "Valor Calculado", "Valor Control", "Diferencia"]].copy()
                    tabla["Valor Calculado"] = tabla["Valor Calculado"].apply(lambda x: f"${x:,.0f}".replace(",","."))
                    tabla["Valor Control"]   = tabla["Valor Control"].apply(lambda x: f"${x:,.0f}".replace(",","."))
                    tabla["Diferencia"]      = tabla["Diferencia"].apply(lambda x: f"${x:,.0f}".replace(",","."))
                    st.dataframe(tabla, use_container_width=True, hide_index=True)

                # Botón descarga Excel del log
                def generar_excel_log(df_e):
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Log de errores"

                    cols = ["Archivo","RUT","Validacion","Nombre validacion","Descripcion",
                            "Formula aplicada","Col control CSV","Valor Calculado","Valor Control","Diferencia"]
                    hdr_fill  = PatternFill("solid", fgColor="2d4a7a")
                    hdr_font  = Font(bold=True, color="FFFFFF", size=10)
                    err_fill  = PatternFill("solid", fgColor="FDECEA")
                    num_fill  = PatternFill("solid", fgColor="FFF3F3")
                    thin = Side(style="thin", color="CCCCCC")
                    border = Border(left=thin, right=thin, top=thin, bottom=thin)

                    ws.append(cols)
                    for cell in ws[1]:
                        cell.fill = hdr_fill
                        cell.font = hdr_font
                        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        cell.border = border

                    for _, row in df_e.iterrows():
                        ws.append([row.get(c, "") for c in cols])
                        last = ws.max_row
                        for i, cell in enumerate(ws[last]):
                            cell.border = border
                            cell.alignment = Alignment(vertical="center", wrap_text=(i in [4,5,6]))
                            if cols[i] in ["Valor Calculado","Valor Control","Diferencia"]:
                                cell.number_format = '#,##0'
                                cell.fill = num_fill
                            else:
                                cell.fill = err_fill

                    # Anchos de columna
                    anchos = [28,18,12,20,45,50,40,18,18,15]
                    for i, ancho in enumerate(anchos, 1):
                        ws.column_dimensions[ws.cell(1, i).column_letter].width = ancho

                    ws.row_dimensions[1].height = 30
                    ws.freeze_panes = "A2"

                    buf = io.BytesIO()
                    wb.save(buf)
                    buf.seek(0)
                    return buf.read()

                excel_log = generar_excel_log(df_err)
                st.download_button(
                    label="⬇️ Descargar log completo (.xlsx)",
                    data=excel_log,
                    file_name=f"log_errores_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        elif st.session_state.validacion_ok and not st.session_state.excel_listo:
            st.markdown("""
            <div class="alert-success">
                ✅ <b>Todas las validaciones se cumplieron correctamente.</b><br>
                Los registros de todos los archivos cuadran sin diferencias.
            </div>""", unsafe_allow_html=True)
            st.markdown("#### ¿Desea generar el archivo de salida?")
            col_a, col_b, col_c, _ = st.columns([1, 1, 1, 4])
            with col_a:
                if st.button("✅ Aceptar"):
                    with st.spinner("Generando archivo de salida..."):
                        df_combined = pd.concat(st.session_state.dfs, ignore_index=True)
                        filas_salida = []
                        todos_multiples = []
                        for fp, grupo in df_combined.groupby("_fecha_proceso"):
                            df_out, multiples = generar_filas_salida(grupo, fp, refs)
                            filas_salida.append(df_out)
                            todos_multiples.extend(multiples)
                        df_final = pd.concat(filas_salida, ignore_index=True) if filas_salida else pd.DataFrame()
                        st.session_state.excel_bytes = generar_excel(df_final)
                        st.session_state.excel_listo = True
                        st.session_state.ruts_multiples = list(set(todos_multiples))
                    st.rerun()
            with col_b:
                if st.button("✖ Cancelar"):
                    for k in ["validacion_ejecutada", "validacion_ok", "errores_validacion", "dfs", "excel_listo", "excel_bytes"]:
                        st.session_state[k] = False if k in ["validacion_ejecutada", "validacion_ok", "excel_listo"] else [] if k in ["errores_validacion", "dfs"] else None
                    st.rerun()
            with col_c:
                if st.button("🚪 Salir"):
                    st.markdown('<div class="alert-warning">Puedes cerrar esta ventana.</div>', unsafe_allow_html=True)
                    st.stop()

        if st.session_state.excel_listo and st.session_state.excel_bytes:
            st.markdown('<div class="alert-success">✅ Archivo generado exitosamente. Haz clic para descargar.</div>', unsafe_allow_html=True)
            nombre_default = f"migracion_{st.session_state.nombre_empresa}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            nombre_usuario = st.text_input("Nombre del archivo", value=nombre_default, help="Puedes personalizar el nombre. Se guardará siempre como .xlsx")
            nombre_final = (nombre_usuario.strip() or nombre_default).replace(".xlsx", "") + ".xlsx"
            st.download_button(
                label="⬇️ Descargar archivo de salida (.xlsx)",
                data=st.session_state.excel_bytes,
                file_name=nombre_final,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
