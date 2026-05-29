import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import calendar
from procesador import procesar_liquidaciones
from generador import generar_archivo_entrada
from mantencion import render_mantencion

st.set_page_config(
    page_title="Liquidaciones Detalle",
    page_icon="📊",
    layout="wide"
)

# ── Estilos ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

.stApp { background-color: #F0F4FA; color: #1E3A5F; }

.header-bar {
    background: #1565C0;
    border-left: 4px solid #00BFA5;
    padding: 20px 28px;
    margin-bottom: 32px;
    border-radius: 0 8px 8px 0;
}

.header-bar h1 {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 1.6rem;
    color: #ffffff;
    margin: 0;
    letter-spacing: -0.5px;
}

.header-bar p {
    color: #B3D1F5;
    margin: 4px 0 0 0;
    font-size: 0.85rem;
}

.module-card {
    background: #ffffff;
    border: 1px solid #D8E4F0;
    border-radius: 12px;
    padding: 28px;
    height: 100%;
    transition: border-color 0.2s;
}

.module-card:hover { border-color: #1565C044; }

.module-title {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.75rem;
    color: #1565C0;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 8px;
}

.module-desc {
    color: #5A6E8A;
    font-size: 0.88rem;
    line-height: 1.6;
    margin-bottom: 20px;
}

.stButton > button {
    background: #1565C0;
    color: #ffffff;
    border: none;
    font-family: 'IBM Plex Mono', monospace;
    font-weight: 600;
    font-size: 0.82rem;
    letter-spacing: 1px;
    padding: 10px 24px;
    border-radius: 6px;
    width: 100%;
    transition: all 0.2s;
}

.stButton > button:hover {
    background: #0D47A1;
    transform: translateY(-1px);
}

.stFileUploader {
    background: #EEF3FA;
    border: 1px dashed #B0C4DE;
    border-radius: 8px;
}

.stFileUploader label { color: #5A6E8A !important; font-size: 0.85rem; }

.success-box {
    background: #E6F4EA;
    border: 1px solid #34A85344;
    border-radius: 8px;
    padding: 16px 20px;
    color: #1B6B3A;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.82rem;
}

.error-box {
    background: #FDECEA;
    border: 1px solid #E5393544;
    border-radius: 8px;
    padding: 16px 20px;
    color: #B71C1C;
    font-size: 0.85rem;
}

.info-box {
    background: #E8F0FB;
    border: 1px solid #1565C044;
    border-radius: 8px;
    padding: 16px 20px;
    color: #1565C0;
    font-size: 0.85rem;
}

div[data-testid="stFileUploadDropzone"] {
    background: #EEF3FA !important;
}

.step-label {
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.7rem;
    color: #7A8FA8;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-bottom: 4px;
}

hr { border-color: #D8E4F0; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background: #ffffff;
    border-bottom: 2px solid #D8E4F0;
}

.stTabs [data-baseweb="tab"] {
    color: #5A6E8A;
    font-family: 'IBM Plex Sans', sans-serif;
}

.stTabs [aria-selected="true"] {
    color: #1565C0 !important;
    border-bottom: 2px solid #1565C0 !important;
}

/* Sidebar si existe */
section[data-testid="stSidebar"] {
    background-color: #1E3A5F;
}

/* Inputs */
input, .stTextInput input {
    background: #ffffff !important;
    border-color: #C8D4E3 !important;
    color: #1E3A5F !important;
}

/* Spinner */
.stSpinner > div { border-top-color: #1565C0 !important; }

</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <h1>⬡ LIQUIDACIONES DETALLE</h1>
    <p>Sistema de procesamiento de remuneraciones · v1.0</p>
</div>
""", unsafe_allow_html=True)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📄 Generar Archivo de Entrada", "⚙️ Procesar Liquidaciones", "🗃️ Mantención de Tablas"])

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — GENERAR ARCHIVO DE ENTRADA
# ═══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("""
    <div class="module-title">Módulo 01</div>
    <div class="module-desc">
        Sube la lista de conceptos y el listado de empresas para generar el archivo Excel
        de entrada con todas las columnas y validaciones configuradas, listo para enviar al cliente.
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="step-label">Paso 1 — Lista de conceptos</div>', unsafe_allow_html=True)
        file_conceptos = st.file_uploader("Lista_de_conceptos.xlsx", type=['xlsx'], key="conceptos")

    with col2:
        st.markdown('<div class="step-label">Paso 2 — Listado de empresas</div>', unsafe_allow_html=True)
        file_empresas = st.file_uploader("listado_empresas.xlsx", type=['xlsx'], key="empresas")

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("GENERAR ARCHIVO DE ENTRADA", key="btn_generar"):
        archivos = [file_conceptos, file_empresas]
        nombres  = ["Lista de conceptos", "Listado de empresas"]
        faltantes = [n for f, n in zip(archivos, nombres) if f is None]

        if faltantes:
            st.markdown(f'<div class="error-box">⚠️ Faltan archivos: {", ".join(faltantes)}</div>', unsafe_allow_html=True)
        else:
            with st.spinner("Generando archivo..."):
                try:
                    output = generar_archivo_entrada(file_conceptos, file_empresas)
                    st.markdown('<div class="success-box">✓ Archivo generado exitosamente</div>', unsafe_allow_html=True)
                    st.download_button(
                        label="⬇ DESCARGAR ARCHIVO DE ENTRADA",
                        data=output,
                        file_name="archivo_entrada_liquidaciones.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.markdown(f'<div class="error-box">Error: {str(e)}</div>', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — PROCESAR LIQUIDACIONES
# ═══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("""
    <div class="module-title">Módulo 02</div>
    <div class="module-desc">
        Sube el archivo de entrada completado por el cliente junto con los archivos de parámetros
        para generar el archivo de liquidaciones detalladas.
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="step-label">Archivo de entrada (completado por cliente)</div>', unsafe_allow_html=True)
        file_entrada = st.file_uploader("archivo_entrada_liquidaciones.xlsx", type=['xlsx'], key="entrada")

    with col2:
        st.markdown('<div class="step-label">Listado de empleados</div>', unsafe_allow_html=True)
        file_empleados2 = st.file_uploader("listado_empleados.xlsx", type=['xlsx'], key="empleados2")

    col3, col4 = st.columns(2)

    with col3:
        st.markdown('<div class="step-label">Listado de empresas</div>', unsafe_allow_html=True)
        file_empresas2 = st.file_uploader("listado_empresas.xlsx", type=['xlsx'], key="empresas2")

    with col4:
        st.markdown('<div class="step-label">Lista de conceptos</div>', unsafe_allow_html=True)
        file_conceptos2 = st.file_uploader("Lista_de_conceptos.xlsx", type=['xlsx'], key="conceptos2")

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("PROCESAR LIQUIDACIONES", key="btn_procesar"):
        archivos = [file_entrada, file_empleados2, file_empresas2, file_conceptos2]
        nombres  = ["Archivo entrada", "Listado empleados", "Listado empresas", "Lista de conceptos"]
        faltantes = [n for f, n in zip(archivos, nombres) if f is None]

        if faltantes:
            st.markdown(f'<div class="error-box">⚠️ Faltan archivos: {", ".join(faltantes)}</div>', unsafe_allow_html=True)
        else:
            with st.spinner("Procesando liquidaciones..."):
                try:
                    output, n_filas, n_trabajadores, sin_empleado, log_bytes = procesar_liquidaciones(
                        file_entrada, file_empleados2, file_empresas2, file_conceptos2
                    )

                    if not sin_empleado:
                        st.markdown(f'''
                        <div class="success-box">
                            ✓ Todos los datos del archivo de entrada se procesaron exitosamente<br>
                            &nbsp;&nbsp;· {n_trabajadores} trabajadores procesados<br>
                            &nbsp;&nbsp;· {n_filas} filas generadas
                        </div>
                        ''', unsafe_allow_html=True)
                    else:
                        st.markdown(f'''
                        <div class="info-box">
                            ⚠️ Se procesaron los datos, pero se detectaron {len(sin_empleado)} registro(s) del archivo de entrada
                            que no están en la lista de empleados.<br>
                            &nbsp;&nbsp;· {n_trabajadores} trabajadores procesados<br>
                            &nbsp;&nbsp;· {n_filas} filas generadas
                        </div>
                        ''', unsafe_allow_html=True)
                        st.download_button(
                            label="⬇ DESCARGAR LOG DE RUTs NO ENCONTRADOS",
                            data=log_bytes,
                            file_name="log_ruts_no_encontrados.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                    st.download_button(
                        label="⬇ DESCARGAR LIQUIDACIONES DETALLADAS",
                        data=output,
                        file_name="liquidaciones_detalladas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.markdown(f'<div class="error-box">Error al procesar: {str(e)}</div>', unsafe_allow_html=True)
                    st.exception(e)

# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — MANTENCIÓN DE TABLAS
# ═══════════════════════════════════════════════════════════════════════════════
with tab3:
    render_mantencion()
