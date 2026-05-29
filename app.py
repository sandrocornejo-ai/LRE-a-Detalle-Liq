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
    .stApp { background-color: #0f1117; color: #e8e8e8; }
    
    .header-bar {
        background: linear-gradient(90deg, #1a1f2e 0%, #0f1117 100%);
        border-left: 4px solid #00d4aa;
        padding: 20px 28px;
        margin-bottom: 32px;
        border-radius: 0 8px 8px 0;
    }
    .header-bar h1 {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.6rem;
        color: #00d4aa;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .header-bar p {
        color: #888;
        margin: 4px 0 0 0;
        font-size: 0.85rem;
    }
    
    .module-card {
        background: #1a1f2e;
        border: 1px solid #2a2f3e;
        border-radius: 12px;
        padding: 28px;
        height: 100%;
        transition: border-color 0.2s;
    }
    .module-card:hover { border-color: #00d4aa44; }
    .module-title {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.75rem;
        color: #00d4aa;
        letter-spacing: 2px;
        text-transform: uppercase;
        margin-bottom: 8px;
    }
    .module-desc {
        color: #aaa;
        font-size: 0.88rem;
        line-height: 1.6;
        margin-bottom: 20px;
    }

    .stButton > button {
        background: #00d4aa;
        color: #0f1117;
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
        background: #00efc0;
        transform: translateY(-1px);
    }

    .stFileUploader {
        background: #12161f;
        border: 1px dashed #2a2f3e;
        border-radius: 8px;
    }
    .stFileUploader label { color: #888 !important; font-size: 0.85rem; }

    .success-box {
        background: #0d2b22;
        border: 1px solid #00d4aa44;
        border-radius: 8px;
        padding: 16px 20px;
        color: #00d4aa;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.82rem;
    }
    .error-box {
        background: #2b0d0d;
        border: 1px solid #ff444444;
        border-radius: 8px;
        padding: 16px 20px;
        color: #ff6b6b;
        font-size: 0.85rem;
    }
    .info-box {
        background: #0d1a2b;
        border: 1px solid #3a6ea544;
        border-radius: 8px;
        padding: 16px 20px;
        color: #7ab8f5;
        font-size: 0.85rem;
    }
    div[data-testid="stFileUploadDropzone"] {
        background: #12161f !important;
    }
    .step-label {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.7rem;
        color: #555;
        letter-spacing: 1px;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    hr { border-color: #2a2f3e; }
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
        nombres = ["Lista de conceptos", "Listado de empresas"]
        faltantes = [n for f, n in zip(archivos, nombres) if f is None]
        if faltantes:
            st.markdown(f'<div class="error-box">⚠️ Faltan archivos: {", ".join(faltantes)}</div>', unsafe_allow_html=True)
        else:
            with st.spinner("Generando archivo..."):
                try:
                    output = generar_archivo_entrada(
                        file_conceptos, file_empresas
                    )
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
        nombres = ["Archivo entrada", "Listado empleados", "Listado empresas", "Lista de conceptos"]
        faltantes = [n for f, n in zip(archivos, nombres) if f is None]
        if faltantes:
            st.markdown(f'<div class="error-box">⚠️ Faltan archivos: {", ".join(faltantes)}</div>', unsafe_allow_html=True)
        else:
            with st.spinner("Procesando liquidaciones..."):
                try:
                    output, n_filas, n_trabajadores = procesar_liquidaciones(
                        file_entrada, file_empleados2, file_empresas2,
                        file_conceptos2
                    )
                    st.markdown(f'''
                    <div class="success-box">
                        ✓ Proceso completado<br>
                        &nbsp;&nbsp;· {n_trabajadores} trabajadores procesados<br>
                        &nbsp;&nbsp;· {n_filas} filas generadas
                    </div>
                    ''', unsafe_allow_html=True)
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
