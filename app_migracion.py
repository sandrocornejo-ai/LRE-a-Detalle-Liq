import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import io
from datetime import datetime
import calendar
from modulo_dt import render_modulo_dt

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
    background-color: #1a2744;
    padding: 14px 28px;
    border-radius: 10px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin-bottom: 28px;
}
.rex-logo {
    background: white;
    color: #1a2744;
    font-weight: 800;
    font-size: 15px;
    padding: 5px 10px;
    border-radius: 6px;
    letter-spacing: 0.5px;
}
.rex-logo span { color: #00b4d8; }
.rex-title { color: white; font-size: 18px; font-weight: 600; margin-left: 16px; }
.rex-badge {
    background: #00b4d8;
    color: white;
    font-size: 11px;
    font-weight: 700;
    padding: 4px 12px;
    border-radius: 20px;
    letter-spacing: 1px;
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

/* Sección parámetros */
.param-title {
    font-size: 0.8rem;
    font-weight: 700;
    color: #1a2744;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    margin-bottom: 1rem;
    padding-bottom: 0.5rem;
    border-bottom: 2px solid #00b4d8;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# EQUIVALENCIAS CONCEPTOS ARCHIVO BASE LRE REX+
# Fuente: equiv_conceptos_ab_rex.xlsx — actualizar aquí si cambia
# Formato: { "Nombre columna entrada": ("Id Concepto", "Tipo") }
# ─────────────────────────────────────────────
EQUIV_CONCEPTOS_REX = {
    "Nro días de licencia médica":                                                                              ("licenciaDias",         "Dato"),
    "Sueldo(2101)":                                                                                             ("sueldoBase",           "Haber afecto"),
    "Sobresueldo(2102)":                                                                                        ("sobresueldoMi",        "Haber afecto"),
    "Comisiones(2103)":                                                                                         ("comisionMi",           "Haber afecto"),
    "Semana corrida(2104)":                                                                                     ("semanaCorr",           "Haber afecto"),
    "Participación(2105)":                                                                                      ("participacionMi",      "Haber afecto"),
    "Gratificación(2106)":                                                                                      ("gratificacion",        "Haber afecto"),
    "Recargo 30% día domingo (Art. 38) (2107)":                                                                 ("recargoMi",            "Haber afecto"),
    "Remuneración variable pagada en vacaciones (Art 71) (cód 2108)":                                           ("remvarMi",             "Haber afecto"),
    "Aguinaldo(2110)":                                                                                          ("AguinaldoMi",          "Haber afecto"),
    "Bonos u otras remun. fijas mensuales(2111)":                                                               ("BonoMi",               "Haber afecto"),
    "Tratos (mensual) (cód 2112)":                                                                              ("tratoMi",              "Haber afecto"),
    "Bonos u otras remuneraciones variables mensuales o superiores a un mes (cód 2113)":                        ("bonsupMi",             "Haber afecto"),
    "Beneficios en especie constitutivos de remuneración (cód 2115)":                                           ("benefMi",              "Haber afecto"),
    "Otras remuneraciones superiores a un mes (cód 2123)":                                                      ("otremMi",              "Haber afecto"),
    "Pago por horas de trabajo sindical (cód 2124)":                                                            ("pagosindMi",           "Haber afecto"),
    "Subsidio por incapacidad laboral por licencia médica(2201)":                                               ("otrosHaberes",         "Haber afecto"),
    "Beca de estudio (Art. 17 N°18 LIR) (cód 2202)":                                                           ("becaMi",               "Haber afecto"),
    "Otros ingresos no constitutivos de renta (Art 17 N°29 LIR) (cód 2204)":                                    ("otingrMi",             "Haber exento"),
    "Colación(2301)":                                                                                           ("colacion",             "Haber exento"),
    "Movilización(2302)":                                                                                       ("movilizacion",         "Haber exento"),
    "Viáticos totales mensual (Art 41) (cód 2303)":                                                             ("gastoMi",              "Haber exento"),
    "Asignación de pérdida de caja(2304)":                                                                      ("asigPerdCajaMi",       "Haber exento"),
    "Asignación de desgaste herramienta(2305)":                                                                 ("AsigDesgHerrMi",       "Haber exento"),
    "Gastos por causa del trabajo (Art 41 CdT) y gastos de representación (Art. 42 Nº1 LIR) (cód 2306)":       ("gastoMi",              "Haber exento"),
    "Sala cuna (Art 203) (cód 2308)":                                                                           ("salaCMi",              "Haber exento"),
    "Asignación familiar legal(2311)":                                                                          ("cargasSimp",           "Haber exento"),
    "Asignación trabajo a distancia o teletrabajo(2309)":                                                       ("AsigTeletrabajoMi",    "Haber exento"),
    "Alojamiento por razones de trabajo (2310)":                                                                ("alojamientoTrabajoMi", "Haber exento"),
    "Asignación de traslación(2312)":                                                                           ("TraslacionMi",         "Haber exento"),
    "Indemnización por feriado legal(2313)":                                                                    ("iasVacaciones",        "Haber exento"),
    "Indemnización años de servicio(2314)":                                                                     ("iasLegal",             "Haber exento"),
    "Indemnización sustitutiva del aviso previo(2315)":                                                         ("iasMes",               "Haber exento"),
    "Indemnización fuero maternal (Art 163 bis) (cód 2316)":                                                    ("indemMMi",             "Haber exento"),
    "Indemnización a todo evento (Art.164) (cód 2331)":                                                         ("indemNMi",             "Haber exento"),
    "Indemnizaciones voluntarias tributables (cód 2417)":                                                       ("iasVoluntaria",        "Haber exento"),
    "Indemnizaciones contractuales tributables (cód 2418)":                                                     ("iasafecta",            "Haber exento"),
    "Cotización obligatoria previsional (AFP o IPS)(3141)":                                                     ("afp",                  "Descuento legal"),
    "Cotización obligatoria salud 7%(3143)":                                                                    ("isapre",               "Descuento legal"),
    "Cotización voluntaria para salud(3144)":                                                                   ("isapre",               "Descuento legal"),
    "Cotización AFC - trabajador(3151)":                                                                        ("cesEmpleado",          "Descuento legal"),
    "Cotización adicional trabajo pesado- trabajador (cód 3154)":                                               ("trabajoPesaEmpl",      "Descuento legal"),
    "Cotización APVi Mod A(3155)":                                                                              ("apvi",                 "Descuento legal"),
    "Cotización APVi Mod B hasta UF50(3156)":                                                                   ("apvi",                 "Descuento legal"),
    "Impuesto retenido por remuneraciones(3161)":                                                               ("impuesto",             "Descuento legal"),
    "Impuesto retenido por indemnizaciones (cód 3162)":                                                         ("imptoindMi",           "Descuento legal"),
    "Mayor retención de impuesto solicitada por el trabajador (cód 3163)":                                      ("ppm",                  "Descuento legal"),
    "Impuesto retenido por reliquidación de remuneraciones devengadas en otros períodos mensuales (cód 3164)":  ("reliquidaImpuesto",    "Descuento legal"),
    "Retención préstamo clase media 2020 (Ley 21.252) (3166)":                                                  ("solidarioremu",        "Descuento legal"),
    "Cuota sindical 1(3171)":                                                                                   ("CuotaSindMi",          "Otros descuentos"),
    "Crédito social CCAF(3110)":                                                                                ("cajaCred",             "Otros descuentos"),
    "Cuota vivienda o educación Art. 58 (cód 3181)":                                                            ("cuotavivMi",           "Otros descuentos"),
    "Crédito cooperativas de ahorro (Art 54 Ley Coop.) (cód 3182)":                                             ("credMi",               "Otros descuentos"),
    "Otros descuentos autorizados y solicitados por el trabajador (cód 3183)":                                  ("otrdescMi",            "Otros descuentos"),
    "Otros descuentos(3185)":                                                                                   ("otrosDesctoMi",        "Otros descuentos"),
    "Pensiones de alimentos(3186)":                                                                             ("retencionJudicial",    "Otros descuentos"),
    "Descuentos por anticipos y préstamos(3188)":                                                               ("AnticipoPrestamoMi",   "Otros descuentos"),
    "AFC - Aporte empleador solidario":                                                                         ("cesAporteSol",         "Aportes empleador"),
    "AFC - Aporte empleador individual":                                                                        ("cesAporteCi",          "Aportes empleador"),
    "Aporte empleador seguro accidentes del trabajo y Ley SANNA(4152)":                                         ("mutual",               "Aportes empleador"),
    "Aporte adicional trabajo pesado- empleador (cód 4154)":                                                    ("trabajoPesa",          "Aportes empleador"),
    "Rebaja zona extrema DL 889 (3167)":                                                                        ("zonaExtrema",          "Aportes empleador"),
    "Aporte empleador seguro invalidez y sobrevivencia(4155)":                                                  ("sis",                  "Aportes empleador"),
    "Total líquido(5501)":                                                                                      ("totalesEmpl",          "Liquido"),
}


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

COLS_OTROS_DESCUENTOS = [
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
CONCEPTOS_LICENCIA_MES_COMPLETO = {
    "sueldoBase", "gratificacion", "afp", "isapre",
    "cesEmpleado", "impuesto", "cesAporteCi", "mutual",
    "sis", "totalesEmpl"
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
# CONSTANTES PARÁMETROS MENSUALES
# ─────────────────────────────────────────────
ARCHIVO_PARAMS = os.path.join(DATA_DIR, "parametrosMesuales.xlsx")
HOJA_PARAMS = "Hoja2"
LABELS_PARAMS = {
    "mes_Proc":           "Mes de proceso (aaaa-mm)",
    "uf_Mes":             "UF del mes ($)",
    "topeImp_Uf_afp":     "Tope imponible AFP (UF)",
    "topeImp_pesos_afp":  "Tope imponible AFP ($)",
    "topeCes_Uf":         "Tope cesantía (UF)",
    "topeCes_pesos":      "Tope cesantía ($)",
    "sis":                "SIS (%)",
    "factor_sis":         "Factor SIS (decimal)",
    "topeSalud_Uf":       "Tope salud (UF)",
    "topeSalud_pesos":    "Tope salud ($)",
    "imm":                "IMM ($)",
    "topeGratif":         "Tope gratificación ($)",
    "monto_Utm":          "Monto UTM ($)",
    "ult_Diames":         "Último día del mes",
    "aporte_Ccaf":        "Aporte CCAF (%)",
    "aporte_Fonasa":      "Aporte FONASA (%)",
    "Formato Fecha":      "Fecha formato (dd/mm/aaaa)",
    "Aporte AFP":         "Aporte AFP (%)",
    "Seg Social Exp vida":"Seg. social / Exp. vida (%)",
}

# ─────────────────────────────────────────────
# FUNCIONES DE CARGA DE REFERENCIAS
# ─────────────────────────────────────────────
@st.cache_data
def cargar_referencias():
    refs = {}
    archivos = {
        "listado_empresas": "listado_empresas.xlsx",
        "inst_mutuales": "inst_mutuales.xlsx",
        "inst_cajas": "inst_cajas.xlsx",
        "inst_afp": "inst_afp.xlsx",
        "inst_salud": "inst_salud.xlsx",
        "cot_afp_hist": "cot_afp_hist.xlsx",
        "parametros": "parametrosMesuales.xlsx",
    }
    errores = []
    for key, fname in archivos.items():
        path = os.path.join(DATA_DIR, fname)
        if os.path.exists(path):
            refs[key] = pd.read_excel(path)
        else:
            errores.append(fname)
    return refs, errores

@st.cache_data(ttl=0)
def cargar_params():
    df = pd.read_excel(ARCHIVO_PARAMS, sheet_name=HOJA_PARAMS, dtype={"mes_Proc": str})
    df["mes_Proc"] = df["mes_Proc"].astype(str).str.strip()
    return df

def guardar_params(df: pd.DataFrame):
    wb = load_workbook(ARCHIVO_PARAMS)
    ws = wb[HOJA_PARAMS]
    ws.delete_rows(2, ws.max_row)
    for _, row in df.iterrows():
        ws.append(list(row))
    wb.save(ARCHIVO_PARAMS)

def render_parametros():
    st.markdown('<div class="param-title">📅 Gestión de parámetros mensuales</div>', unsafe_allow_html=True)
    if not os.path.exists(ARCHIVO_PARAMS):
        st.error(f"⚠️ No se encontró el archivo `{ARCHIVO_PARAMS}`.")
        return
    df_p = cargar_params()
    tab_add, tab_edit, tab_view = st.tabs(["➕ Agregar mes", "✏️ Editar mes", "📊 Ver tabla"])

    with tab_add:
        ultimo = df_p["mes_Proc"].dropna().iloc[-1] if not df_p.empty else "2026-01"
        try:
            dt_ult = datetime.strptime(str(ultimo)[:7], "%Y-%m")
            mes_sig = f"{dt_ult.year + 1}-01" if dt_ult.month == 12 else f"{dt_ult.year}-{dt_ult.month + 1:02d}"
        except Exception:
            mes_sig = ""
        nuevo_mes = st.text_input("Mes de proceso", value=mes_sig, placeholder="aaaa-mm", key="pm_nuevo_mes")
        mes_ok = False
        if nuevo_mes:
            try:
                dt = datetime.strptime(nuevo_mes[:7], "%Y-%m")
                ult_dia = calendar.monthrange(dt.year, dt.month)[1]
                mes_ok = True
                if nuevo_mes in df_p["mes_Proc"].values:
                    st.warning(f"⚠️ El mes **{nuevo_mes}** ya existe. Usa **Editar mes** para modificarlo.")
                    mes_ok = False
            except ValueError:
                st.error("Formato inválido. Usa aaaa-mm (ej: 2026-07)")
                ult_dia = 30
        if mes_ok:
            ult_row = df_p.iloc[-1] if not df_p.empty else {}
            def vref(c):
                try:
                    v = ult_row.get(c, 0)
                    return float(v) if pd.notna(v) else 0.0
                except Exception:
                    return 0.0
            nuevo = {"mes_Proc": nuevo_mes}
            campos = [c for c in LABELS_PARAMS if c not in ("mes_Proc", "Formato Fecha", "factor_sis", "ult_Diames")]
            cols_f = st.columns(3)
            for i, col in enumerate(campos):
                with cols_f[i % 3]:
                    fmt = "%.4f" if col in ("sis", "Aporte AFP", "Seg Social Exp vida", "aporte_Ccaf", "aporte_Fonasa") else "%.2f"
                    nuevo[col] = st.number_input(LABELS_PARAMS[col], value=vref(col), format=fmt, key=f"pm_new_{col}")
            nuevo["factor_sis"] = round(nuevo.get("sis", 0) / 100, 6)
            nuevo["ult_Diames"] = ult_dia
            nuevo["Formato Fecha"] = f"{ult_dia:02d}/{dt.month:02d}/{dt.year}"
            st.caption(f"📌 `factor_sis` = {nuevo['factor_sis']} | Último día del mes = {ult_dia}")
            if st.button("💾 Guardar nuevo mes", key="pm_btn_add"):
                fila = {col: nuevo.get(col) for col in df_p.columns}
                df_nuevo = pd.concat([df_p, pd.DataFrame([fila])], ignore_index=True)
                try:
                    guardar_params(df_nuevo)
                    st.cache_data.clear()
                    st.success(f"✅ Mes **{nuevo_mes}** agregado.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al guardar: {e}")

    with tab_edit:
        meses = list(reversed(df_p["mes_Proc"].dropna().tolist()))
        mes_sel = st.selectbox("Mes a editar", meses, key="pm_mes_sel")
        if mes_sel:
            idx = df_p[df_p["mes_Proc"] == mes_sel].index[0]
            fila = df_p.loc[idx].copy()
            editado = {"mes_Proc": mes_sel}
            campos_e = [c for c in LABELS_PARAMS if c not in ("mes_Proc", "Formato Fecha", "factor_sis", "ult_Diames")]
            cols_e = st.columns(3)
            for i, col in enumerate(campos_e):
                with cols_e[i % 3]:
                    try:
                        val = float(fila.get(col, 0)) if pd.notna(fila.get(col)) else 0.0
                    except Exception:
                        val = 0.0
                    fmt = "%.4f" if col in ("sis", "Aporte AFP", "Seg Social Exp vida", "aporte_Ccaf", "aporte_Fonasa") else "%.2f"
                    editado[col] = st.number_input(LABELS_PARAMS[col], value=val, format=fmt, key=f"pm_edit_{col}")
            editado["factor_sis"] = round(editado.get("sis", 0) / 100, 6)
            try:
                dt_e = datetime.strptime(mes_sel[:7], "%Y-%m")
                ult_dia_e = calendar.monthrange(dt_e.year, dt_e.month)[1]
                editado["ult_Diames"] = ult_dia_e
                editado["Formato Fecha"] = f"{ult_dia_e:02d}/{dt_e.month:02d}/{dt_e.year}"
            except Exception:
                pass
            st.caption(f"📌 `factor_sis` = {editado['factor_sis']}")
            if st.button("💾 Guardar cambios", key="pm_btn_edit"):
                for col in df_p.columns:
                    df_p.at[idx, col] = editado.get(col, df_p.at[idx, col])
                try:
                    guardar_params(df_p)
                    st.cache_data.clear()
                    st.success(f"✅ Mes **{mes_sel}** actualizado.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al guardar: {e}")

    with tab_view:
        col_f, _ = st.columns([1, 3])
        with col_f:
            filtro = st.text_input("🔍 Filtrar por año", placeholder="ej: 2026", key="pm_filtro")
        df_v = df_p[df_p["mes_Proc"].str.startswith(filtro)] if filtro else df_p.copy()
        st.dataframe(df_v.rename(columns=LABELS_PARAMS), use_container_width=True, hide_index=True, height=480)
        st.caption(f"Total: {len(df_v)} períodos")

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

def safe_num(v, default=0):
    """Convierte un valor a número seguro."""
    try:
        v = float(str(v).replace(",", ".").strip())
        return v if pd.notna(v) else default
    except Exception:
        return default

COD_COL_REX = {
    1116: "Nro días de licencia médica",
    2101: "Sueldo(2101)", 2102: "Sobresueldo(2102)", 2103: "Comisiones(2103)",
    2104: "Semana corrida(2104)", 2105: "Participación(2105)", 2106: "Gratificación(2106)",
    2107: "Recargo 30% día domingo (Art. 38) (2107)",
    2108: "Remuneración variable pagada en vacaciones (Art 71) (cód 2108)",
    2110: "Aguinaldo(2110)", 2111: "Bonos u otras remun. fijas mensuales(2111)",
    2112: "Tratos (mensual) (cód 2112)",
    2113: "Bonos u otras remuneraciones variables mensuales o superiores a un mes (cód 2113)",
    2115: "Beneficios en especie constitutivos de remuneración (cód 2115)",
    2123: "Otras remuneraciones superiores a un mes (cód 2123)",
    2124: "Pago por horas de trabajo sindical (cód 2124)",
    2201: "Subsidio por incapacidad laboral por licencia médica(2201)",
    2202: "Beca de estudio (Art. 17 N°18 LIR) (cód 2202)",
    2204: "Otros ingresos no constitutivos de renta (Art 17 N°29 LIR) (cód 2204)",
    2301: "Colación(2301)", 2302: "Movilización(2302)",
    2303: "Viáticos totales mensual (Art 41) (cód 2303)",
    2304: "Asignación de pérdida de caja(2304)",
    2305: "Asignación de desgaste herramienta(2305)",
    2306: "Gastos por causa del trabajo (Art 41 CdT) y gastos de representación (Art. 42 Nº1 LIR) (cód 2306)",
    2308: "Sala cuna (Art 203) (cód 2308)",
    2309: "Asignación trabajo a distancia o teletrabajo(2309)",
    2310: "Alojamiento por razones de trabajo (2310)",
    2311: "Asignación familiar legal(2311)",
    2312: "Asignación de traslación(2312)",
    2313: "Indemnización por feriado legal(2313)",
    2314: "Indemnización años de servicio(2314)",
    2315: "Indemnización sustitutiva del aviso previo(2315)",
    2316: "Indemnización fuero maternal (Art 163 bis) (cód 2316)",
    2331: "Indemnización a todo evento (Art.164) (cód 2331)",
    2417: "Indemnizaciones voluntarias tributables (cód 2417)",
    2418: "Indemnizaciones contractuales tributables (cód 2418)",
    3141: "Cotización obligatoria previsional (AFP o IPS)(3141)",
    3143: "Cotización obligatoria salud 7%(3143)",
    3144: "Cotización voluntaria para salud(3144)",
    3151: "Cotización AFC - trabajador(3151)",
    3154: "Cotización adicional trabajo pesado- trabajador (cód 3154)",
    3155: "Cotización APVi Mod A(3155)", 3156: "Cotización APVi Mod B hasta UF50(3156)",
    3161: "Impuesto retenido por remuneraciones(3161)",
    3162: "Impuesto retenido por indemnizaciones (cód 3162)",
    3163: "Mayor retención de impuesto solicitada por el trabajador (cód 3163)",
    3164: "Impuesto retenido por reliquidación de remuneraciones devengadas en otros períodos mensuales (cód 3164)",
    3166: "Retención préstamo clase media 2020 (Ley 21.252) (3166)",
    3171: "Cuota sindical 1(3171)", 3110: "Crédito social CCAF(3110)",
    3181: "Cuota vivienda o educación Art. 58 (cód 3181)",
    3182: "Crédito cooperativas de ahorro (Art 54 Ley Coop.) (cód 3182)",
    3183: "Otros descuentos autorizados y solicitados por el trabajador (cód 3183)",
    3185: "Otros descuentos(3185)", 3186: "Pensiones de alimentos(3186)",
    3188: "Descuentos por anticipos y préstamos(3188)",
    3167: "Rebaja zona extrema DL 889 (3167)",
    4151: "AFC - Aporte empleador solidario",
    4152: "Aporte empleador seguro accidentes del trabajo y Ley SANNA(4152)",
    4154: "Aporte adicional trabajo pesado- empleador (cód 4154)",
    4155: "Aporte empleador seguro invalidez y sobrevivencia(4155)",
    5501: "Total líquido(5501)",
}

CODIGOS_HABERES_REX = [
    2101, 2102, 2103, 2104, 2105, 2106, 2107, 2108, 2110, 2111,
    2112, 2113, 2115, 2123, 2124, 2201, 2202, 2204, 2301, 2302,
    2303, 2304, 2305, 2306, 2308, 2311, 2309, 2310, 2312, 2313,
    2314, 2315, 2316, 2331, 2417, 2418
]

CODIGOS_DESCUENTOS_REX = [
    3141, 3143, 3144, 3151, 3154, 3155, 3156, 3161, 3162, 3163,
    3164, 3166, 3171, 3110, 3181, 3182, 3183, 3185, 3186, 3188
]

def get_col_rex(df, codigo):
    nombre = COD_COL_REX.get(codigo)
    if nombre and nombre in df.columns:
        return nombre
    return None

def safe_sum_rex(df, codigos):
    total = pd.Series(0.0, index=df.index)
    for cod in codigos:
        col = get_col_rex(df, cod)
        if col:
            total += pd.to_numeric(df[col], errors="coerce").fillna(0)
    return total

def validar_liquidez_rex(df):
    tol = 1
    col_dias_trab = "Nro días trabajados"
    if col_dias_trab in df.columns:
        df_val = df[pd.to_numeric(df[col_dias_trab], errors="coerce").fillna(0) != 0].copy()
    else:
        df_val = df.copy()

    suma_haberes    = safe_sum_rex(df_val, CODIGOS_HABERES_REX)
    suma_descuentos = safe_sum_rex(df_val, CODIGOS_DESCUENTOS_REX)
    liquido_calc    = suma_haberes - suma_descuentos

    col_5501 = get_col_rex(df_val, 5501)
    if not col_5501:
        return pd.DataFrame(), False

    liquido_ctrl = pd.to_numeric(df_val[col_5501], errors="coerce").fillna(0)
    mask = (liquido_calc - liquido_ctrl).abs() > tol

    if not mask.any():
        return pd.DataFrame(), False

    df_errores = df_val[mask].copy()
    df_errores["_Suma haberes calculada"]    = suma_haberes[mask].round(2)
    df_errores["_Suma descuentos calculada"] = suma_descuentos[mask].round(2)
    df_errores["_Liquido calculado"]         = liquido_calc[mask].round(2)
    df_errores["_Liquido archivo (5501)"]    = liquido_ctrl[mask].round(2)
    df_errores["_Diferencia"]                = (liquido_calc - liquido_ctrl)[mask].round(2)

    return df_errores, True


def generar_filas_salida(df, fecha_proceso, refs):
    """Genera las filas del archivo de salida via pivot de conceptos."""
    filas = []

    # Cargar lookups
    empresas  = refs.get("listado_empresas", pd.DataFrame())
    mutuales  = refs.get("inst_mutuales", pd.DataFrame())
    cajas     = refs.get("inst_cajas", pd.DataFrame())
    cot_afp   = refs.get("cot_afp_hist", pd.DataFrame())
    params    = refs.get("parametros", pd.DataFrame())

    # Diccionario de equivalencias desde constante estática
    # { nombre_columna: id_concepto }
    equiv_dict = {col: conc for col, (conc, _) in EQUIV_CONCEPTOS_REX.items()}
    # { id_concepto: [col1, col2, ...] }
    concepto_a_cols = {}
    for col, (conc, _) in EQUIV_CONCEPTOS_REX.items():
        concepto_a_cols.setdefault(conc, []).append(col)
    # { id_concepto: tipo }
    tipo_dict = {conc: tipo for col, (conc, tipo) in EQUIV_CONCEPTOS_REX.items()}

    # Parámetros mensuales
    tope_afp = 0
    tope_ces = 0
    if not params.empty:
        if "topeImp_pesos_afp" in params.columns:
            tope_afp = safe_num(params["topeImp_pesos_afp"].iloc[0])
        if "topeCes_pesos" in params.columns:
            tope_ces = safe_num(params["topeCes_pesos"].iloc[0])

    # Columnas de conceptos presentes en el archivo
    cols_concepto = [c for c in df.columns if c in equiv_dict]
    # Conceptos únicos a generar (en orden de aparición)
    conceptos_unicos = list(dict.fromkeys([equiv_dict[c] for c in cols_concepto]))

    for _, row in df.iterrows():
        rut             = str(row.get("Id empleado", "")).strip()
        empresa_entrada = str(row.get("Id de empresa", "")).strip()
        dias_trabajados = safe_num(row.get("Nro días trabajados", 0))
        dias_licencia   = safe_num(row.get("Nro días de licencia médica", 0))
        dias_vacaciones = safe_num(row.get("Nro días de vacaciones en el mes(1117)", 0))
        sueldo          = safe_num(row.get("Sueldo(2101)", 0))
        rebaja_zona     = safe_num(row.get("Rebaja zona extrema DL 889 (3167)", 0))
        numero_contrato = safe_num(row.get("Número de contrato", 1)) or 1
        jornada         = row.get("jornada", "C") or "C"

        # Institución directamente desde el archivo
        afp_empleado    = str(row.get("afp", "")).strip()
        isapre_empleado = str(row.get("isapre", "")).strip()
        mutual_empleado = str(row.get("Mutual", "")).strip()
        ccaf_empleado   = str(row.get("Ccaf", "")).strip()

        # Lookup empresa
        empresa_salida = empresa_entrada
        if empresa_entrada and not empresas.empty and "Nombre" in empresas.columns:
            emp2 = empresas[empresas["Nombre"].astype(str).str.strip() == empresa_entrada]
            if not emp2.empty:
                empresa_salida = str(emp2.iloc[0].iloc[0]).strip()

        # Total imponible (suma haberes afectos)
        cols_afectos = concepto_a_cols.get("sueldoBase", []) + \
                       [c for conc, cols in concepto_a_cols.items()
                        if conc in GRUPOS_AFP_MUTUAL_AFECTO for c in cols]
        total_imponible = sum(safe_num(row.get(c, 0)) for c in set(cols_afectos) if c in df.columns)

        monto_init = round((sueldo / dias_trabajados) * 30, 0) if dias_trabajados > 0 else 0

        # Conceptos ya generados (para evitar duplicados)
        conceptos_ya_generados = set()

        # Fila por cada concepto único
        for id_concepto in conceptos_unicos:

            # Licencia mes completo: solo conceptos permitidos
            if dias_trabajados == 0 and id_concepto not in CONCEPTOS_LICENCIA_MES_COMPLETO:
                continue

            if id_concepto in conceptos_ya_generados:
                continue

            # Sumar todas las columnas que mapean a este concepto
            cols_del_concepto = [c for c in concepto_a_cols.get(id_concepto, []) if c in df.columns]
            monto = sum(safe_num(row.get(c, 0)) for c in cols_del_concepto)

            # Saltar si monto 0 excepto conceptos que siempre se generan
            if monto == 0 and id_concepto not in {"cesEmpleado", "impuesto"} \
               and not (dias_trabajados == 0 and id_concepto in CONCEPTOS_LICENCIA_MES_COMPLETO):
                continue

            # Id de institución
            id_institucion = ""
            if id_concepto in GRUPOS_AFP:
                id_institucion = afp_empleado
            elif id_concepto in GRUPOS_ISAPRE:
                id_institucion = isapre_empleado
            elif id_concepto in GRUPOS_MUTUAL:
                if not mutuales.empty and "nombre_mutual" in mutuales.columns and "id_mutual" in mutuales.columns:
                    m = mutuales[mutuales["nombre_mutual"].astype(str).str.strip() == mutual_empleado]
                    if not m.empty:
                        id_institucion = m.iloc[0]["id_mutual"]
            elif id_concepto in GRUPOS_CCAF:
                if not cajas.empty and "nombre_ccaf" in cajas.columns and "id_ccaf" in cajas.columns:
                    c = cajas[cajas["nombre_ccaf"].astype(str).str.strip() == ccaf_empleado]
                    if not c.empty:
                        id_institucion = c.iloc[0]["id_ccaf"]

            # Afecto
            afecto = ""
            if id_concepto in GRUPOS_AFP_MUTUAL_AFECTO:
                afecto = min(total_imponible, tope_afp) if tope_afp > 0 else total_imponible
            elif id_concepto in GRUPOS_CES_AFECTO:
                afecto = min(total_imponible, tope_ces) if tope_ces > 0 else total_imponible

            # Cotización de jubilación
            cot_jubilacion = 0
            if id_concepto == "afp":
                key_afp = f"{fecha_proceso}{id_institucion}"
                if not cot_afp.empty and "id_afp_hist" in cot_afp.columns:
                    r = cot_afp[cot_afp["id_afp_hist"] == key_afp]
                    if not r.empty:
                        cot_jubilacion = safe_num(r.iloc[0].get("cot_hist_afp", 0)) * 100
            elif id_concepto == "sis":
                key_sis = f"{fecha_proceso}{id_institucion}"
                if not cot_afp.empty and "id_afp_hist" in cot_afp.columns:
                    r = cot_afp[cot_afp["id_afp_hist"] == key_sis]
                    if not r.empty:
                        cot_jubilacion = safe_num(r.iloc[0].get("sis_hist", 0)) * 100
            elif id_concepto == "cesEmpleado":
                cot_jubilacion = 0.6
            elif id_concepto == "isapre":
                cot_jubilacion = monto
            elif id_concepto == "mutual":
                if not empresas.empty and "Mutual" in empresas.columns and "Cotización Mutual" in empresas.columns:
                    e = empresas[empresas["Mutual"].astype(str).str.strip() == mutual_empleado]
                    if not e.empty:
                        cot_jubilacion = safe_num(e.iloc[0]["Cotización Mutual"])
            elif id_concepto == "licenciaDias":
                cot_jubilacion = dias_licencia

            conceptos_ya_generados.add(id_concepto)

            filas.append({
                "Fecha de proceso":        fecha_proceso,
                "Id empleado":             rut,
                "Número de contrato":      int(numero_contrato),
                "Id del concepto":         id_concepto,
                "Monto del concepto":      monto,
                "Afecto":                  afecto,
                "Id de institución":       id_institucion,
                "Cotización de jubilación": cot_jubilacion,
                "Días de licencias":       dias_licencia,
                "Días trabajados":         dias_vacaciones,
                "Fecha de aplicación":     "x",
                "Empresa":                 empresa_salida,
                "Total de rebajas por L":  rebaja_zona,
                "Jornada":                 jornada,
                "Fase":                    1,
                "Días de vacaciones":      dias_vacaciones,
                "Monto Init":              monto_init,
            })

        # Fila adicional licenciaDias si aplica
        if dias_licencia > 0 and "licenciaDias" not in conceptos_ya_generados:
            filas.append({
                "Fecha de proceso":        fecha_proceso,
                "Id empleado":             rut,
                "Número de contrato":      int(numero_contrato),
                "Id del concepto":         "licenciaDias",
                "Monto del concepto":      dias_licencia,
                "Afecto":                  "",
                "Id de institución":       "",
                "Cotización de jubilación": dias_licencia,
                "Días de licencias":       dias_licencia,
                "Días trabajados":         dias_vacaciones,
                "Fecha de aplicación":     "x",
                "Empresa":                 empresa_salida,
                "Total de rebajas por L":  rebaja_zona,
                "Jornada":                 jornada,
                "Fase":                    1,
                "Días de vacaciones":      dias_vacaciones,
                "Monto Init":              monto_init,
            })

    return pd.DataFrame(filas)


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

# Header
st.markdown("""
<div class="rex-header">
    <div style="display:flex; align-items:center; gap:12px;">
        <div class="rex-logo">Rex<span>+</span></div>
        <span class="rex-title">Liquidaciones en detalle desde LRE</span>
    </div>
    <div class="rex-badge">PRODUCCIÓN</div>
</div>
""", unsafe_allow_html=True)

# ── NAVEGACIÓN PRINCIPAL ──
nav_migracion, nav_dt = st.tabs(["📂 Migración desde archivo base LRE de Rex", "🏛️ Migración DT"])

# Cargar referencias compartidas (disponibles para todos los tabs)
refs, errores_refs = cargar_referencias()
if errores_refs:
    st.markdown(f'<div class="alert-warning">⚠️ Archivos de referencia no encontrados en <b>/data</b>: {", ".join(errores_refs)}</div>', unsafe_allow_html=True)

with nav_migracion:
    st.markdown('<div class="section-title">📂 Migración desde archivo base LRE de Rex</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-sub">Sube el archivo Excel exportado desde Rex+ para generar el archivo de salida listo para importar.</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="step-card">
            <div class="step-label">PASO 1</div>
            <div class="step-title">Subir archivo base LRE</div>
            <div class="step-desc">Archivo Excel exportado desde Rex+ con las liquidaciones del período.</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="step-card">
            <div class="step-label">PASO 2</div>
            <div class="step-title">Subir listado de empleados</div>
            <div class="step-desc">Archivo xlsx con los empleados del período a procesar.</div>
        </div>""", unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="step-card">
            <div class="step-label">PASO 3</div>
            <div class="step-title">Descargar Excel</div>
            <div class="step-desc">Se genera el archivo de salida listo para importar en Rex+.</div>
        </div>""", unsafe_allow_html=True)

    st.markdown('<hr class="rex-divider">', unsafe_allow_html=True)

    # Upload archivo base LRE
    st.markdown("### 📤 Archivo base LRE de Rex+")
    archivo_lre = st.file_uploader(
        "Selecciona el archivo Excel exportado desde Rex+",
        type=["xlsx"],
        accept_multiple_files=False,
        key="lre_upload"
    )

    # Upload listado_empleados
    st.markdown("### 👥 Listado de empleados del período")
    archivo_empleados = st.file_uploader(
        "Sube el archivo listado_empleados.xlsx correspondiente al período a procesar",
        type=["xlsx"],
        accept_multiple_files=False,
        key="emp_upload",
        help="Este archivo cambia en cada proceso."
    )

    if archivo_empleados:
        st.markdown(f'<div class="alert-success">✅ Listado de empleados cargado: <b>{archivo_empleados.name}</b></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="alert-warning">⚠️ Debes subir el listado de empleados del período para ejecutar el proceso.</div>', unsafe_allow_html=True)

    if archivo_lre:
        st.markdown(f'<div class="alert-success">✅ Archivo LRE cargado: <b>{archivo_lre.name}</b></div>', unsafe_allow_html=True)

    if archivo_lre and archivo_empleados:
        if st.button("▶ Generar archivo de salida"):
            with st.spinner("Procesando archivo..."):
                try:
                    archivo_lre.seek(0)
                    df = pd.read_excel(archivo_lre)
                    archivo_empleados.seek(0)
                    refs["listado_empleados"] = pd.read_excel(archivo_empleados)

                    # Extraer fecha de proceso
                    if "Fecha de proceso" in df.columns:
                        fecha_proceso = str(df["Fecha de proceso"].iloc[0])[:7]
                    else:
                        fecha_proceso = datetime.now().strftime("%Y-%m")

                    # ── Validación: 5501 = haberes - descuentos ──
                    df_errores_val, hay_errores = validar_liquidez_rex(df)

                    if hay_errores:
                        st.markdown(f"""
                        <div class="alert-error">
                            ❌ <b>Validación fallida.</b> Se encontraron <b>{len(df_errores_val)} trabajador(es)</b>
                            donde el Total líquido (5501) no coincide con la diferencia entre haberes y descuentos.<br>
                            No se generó el archivo de salida. Descarga el log para revisar las diferencias.
                        </div>""", unsafe_allow_html=True)

                        log_buf = io.BytesIO()
                        df_errores_val.to_excel(log_buf, index=False, sheet_name="Errores validación")
                        log_buf.seek(0)
                        st.download_button(
                            label="⬇️ Descargar log de errores (.xlsx)",
                            data=log_buf,
                            file_name=f"log_validacion_{fecha_proceso}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.markdown('<div class="alert-success">✅ <b>Validación exitosa.</b> Todos los registros cuadran. Generando archivo de salida...</div>', unsafe_allow_html=True)

                        df_final = generar_filas_salida(df, fecha_proceso, refs)

                        if df_final.empty:
                            st.markdown('<div class="alert-error">❌ El archivo de salida quedó vacío. Columnas detectadas en el archivo de entrada:</div>', unsafe_allow_html=True)
                            st.write(list(df.columns))
                        else:
                            excel_bytes = generar_excel(df_final)
                            st.markdown(f'<div class="alert-success">✅ Archivo generado: <b>{len(df_final)} filas</b>.</div>', unsafe_allow_html=True)
                            st.download_button(
                                label="⬇️ Descargar archivo de salida (.xlsx)",
                                data=excel_bytes,
                                file_name=f"migracion_lre_{fecha_proceso}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                except Exception as e:
                    import traceback
                    st.markdown(f'<div class="alert-error">❌ Error: <b>{e}</b></div>', unsafe_allow_html=True)
                    st.code(traceback.format_exc())

with nav_dt:
    render_modulo_dt(refs)
