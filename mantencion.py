import streamlit as st
import pandas as pd
import requests
import json

SUPABASE_URL = st.secrets['SUPABASE_URL']
SUPABASE_KEY = st.secrets['SUPABASE_KEY']

HEADERS = {
    'apikey': SUPABASE_KEY,
    'Authorization': f'Bearer {SUPABASE_KEY}',
    'Content-Type': 'application/json',
    'Prefer': 'return=representation'
}

# ── Configuración de tablas ───────────────────────────────────────────────────
TABLAS = {
    'AFP': {
        'tabla': 'inst_afp',
        'columnas': ['id_afp', 'clas_afp', 'nombre_afp', 'rut_afp', 'cod_prev_afp', 'cot_afp', 'observ_afp'],
        'labels':   ['ID AFP', 'Clasificación', 'Nombre AFP', 'RUT', 'Cód. Previred', 'Cotización %', 'Observación'],
        'tipos':    ['text', 'text', 'text', 'text', 'text', 'number', 'text'],
        'key':      'id_afp'
    },
    'Salud': {
        'tabla': 'inst_salud',
        'columnas': ['id_inst', 'clasif', 'nombre_inst', 'rut_inst', 'equiv_previred'],
        'labels':   ['ID Institución', 'Clasificación', 'Nombre', 'RUT', 'Equiv. Previred'],
        'tipos':    ['text', 'text', 'text', 'text', 'text'],
        'key':      'id_inst'
    },
    'Mutuales': {
        'tabla': 'inst_mutuales',
        'columnas': ['id_institucion', 'clasificacion', 'nombre_institucion', 'doc_identidad', 'codigo_equivalente', 'valor', 'valor2', 'valor3'],
        'labels':   ['ID', 'Clasificación', 'Nombre Institución', 'Doc. Identidad', 'Cód. Equivalente', 'Valor', 'Valor 2', 'Valor 3'],
        'tipos':    ['text', 'text', 'text', 'text', 'text', 'number', 'number', 'number'],
        'key':      'id_institucion'
    },
    'Cajas de Compensación': {
        'tabla': 'inst_cajas',
        'columnas': ['id_institucion', 'clasificacion', 'nombre_institucion', 'doc_identidad', 'codigo_equivalente', 'valor', 'valor2', 'valor3'],
        'labels':   ['ID', 'Clasificación', 'Nombre Institución', 'Doc. Identidad', 'Cód. Equivalente', 'Valor', 'Valor 2', 'Valor 3'],
        'tipos':    ['text', 'text', 'text', 'text', 'text', 'number', 'number', 'number'],
        'key':      'id_institucion'
    },
    'Parámetros Mensuales': {
        'tabla': 'parametros_mensuales',
        'columnas': ['mes_proc', 'uf_mes', 'tope_imp_pesos_afp', 'tope_ces_pesos', 'sis', 'tope_salud_pesos', 'imm', 'tope_gratif', 'monto_utm', 'aporte_ccaf', 'aporte_fonasa', 'aporte_afp', 'seg_social_exp_vida'],
        'labels':   ['Mes Proceso', 'UF Mes', 'Tope Imp. AFP ($)', 'Tope Ces. ($)', 'SIS %', 'Tope Salud ($)', 'IMM', 'Tope Gratif.', 'Monto UTM', 'Aporte CCAF', 'Aporte Fonasa', 'Aporte AFP', 'Seg. Social Exp. Vida'],
        'tipos':    ['text', 'number', 'number', 'number', 'number', 'number', 'number', 'number', 'number', 'number', 'number', 'number', 'number'],
        'key':      'mes_proc'
    },
}

def get_data(tabla):
    url = f'{SUPABASE_URL}/rest/v1/{tabla}?select=*&order=id'
    r = requests.get(url, headers=HEADERS)
    if r.status_code == 200:
        return r.json()
    return []

def insert_record(tabla, data):
    url = f'{SUPABASE_URL}/rest/v1/{tabla}'
    r = requests.post(url, headers=HEADERS, json=data)
    return r.status_code in [200, 201]

def update_record(tabla, record_id, data):
    url = f'{SUPABASE_URL}/rest/v1/{tabla}?id=eq.{record_id}'
    r = requests.patch(url, headers=HEADERS, json=data)
    return r.status_code in [200, 204]

def delete_record(tabla, record_id):
    url = f'{SUPABASE_URL}/rest/v1/{tabla}?id=eq.{record_id}'
    r = requests.delete(url, headers=HEADERS)
    return r.status_code in [200, 204]

def render_mantencion():
    st.markdown("""
    <div class="module-title">Módulo 03</div>
    <div class="module-desc">
        Administra las tablas de referencia del sistema. Agrega, modifica o elimina registros
        de AFP, Salud, Mutuales, Cajas de Compensación y Parámetros Mensuales.
    </div>
    """, unsafe_allow_html=True)

    tabla_sel = st.selectbox("Seleccionar tabla", list(TABLAS.keys()), key="tabla_sel")
    config = TABLAS[tabla_sel]

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Cargar datos ──────────────────────────────────────────────────────────
    data = get_data(config['tabla'])

    if data:
        df = pd.DataFrame(data)
        cols_mostrar = ['id'] + config['columnas']
        cols_mostrar = [c for c in cols_mostrar if c in df.columns]
        df_show = df[cols_mostrar].copy()
        df_show.columns = ['ID DB'] + config['labels'][:len(cols_mostrar)-1]
        st.dataframe(df_show, use_container_width=True, hide_index=True)
    else:
        st.markdown('<div class="info-box">No hay registros en esta tabla.</div>', unsafe_allow_html=True)
        data = []

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Acciones ──────────────────────────────────────────────────────────────
    accion = st.radio("Acción", ["➕ Agregar", "✏️ Modificar", "🗑️ Eliminar"], horizontal=True, key="accion")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── AGREGAR ───────────────────────────────────────────────────────────────
    if accion == "➕ Agregar":
        st.markdown('<div class="step-label">Nuevo registro</div>', unsafe_allow_html=True)
        nuevo = {}
        cols = st.columns(3)
        for i, (col, label, tipo) in enumerate(zip(config['columnas'], config['labels'], config['tipos'])):
            with cols[i % 3]:
                if tipo == 'number':
                    nuevo[col] = st.number_input(label, value=0.0, key=f"new_{col}")
                else:
                    nuevo[col] = st.text_input(label, key=f"new_{col}")

        if st.button("GUARDAR REGISTRO", key="btn_agregar"):
            if insert_record(config['tabla'], nuevo):
                st.markdown('<div class="success-box">✓ Registro agregado exitosamente</div>', unsafe_allow_html=True)
                st.rerun()
            else:
                st.markdown('<div class="error-box">Error al agregar el registro</div>', unsafe_allow_html=True)

    # ── MODIFICAR ─────────────────────────────────────────────────────────────
    elif accion == "✏️ Modificar":
        if not data:
            st.markdown('<div class="info-box">No hay registros para modificar.</div>', unsafe_allow_html=True)
            return

        opciones = {f"ID {r['id']} — {r.get(config['key'], '')}": r['id'] for r in data}
        sel = st.selectbox("Seleccionar registro", list(opciones.keys()), key="sel_mod")
        record_id = opciones[sel]
        registro = next(r for r in data if r['id'] == record_id)

        st.markdown('<div class="step-label">Editar valores</div>', unsafe_allow_html=True)
        editado = {}
        cols = st.columns(3)
        for i, (col, label, tipo) in enumerate(zip(config['columnas'], config['labels'], config['tipos'])):
            with cols[i % 3]:
                val_actual = registro.get(col)
                if tipo == 'number':
                    editado[col] = st.number_input(label, value=float(val_actual) if val_actual else 0.0, key=f"edit_{col}")
                else:
                    editado[col] = st.text_input(label, value=str(val_actual) if val_actual else '', key=f"edit_{col}")

        if st.button("GUARDAR CAMBIOS", key="btn_modificar"):
            if update_record(config['tabla'], record_id, editado):
                st.markdown('<div class="success-box">✓ Registro actualizado exitosamente</div>', unsafe_allow_html=True)
                st.rerun()
            else:
                st.markdown('<div class="error-box">Error al actualizar el registro</div>', unsafe_allow_html=True)

    # ── ELIMINAR ──────────────────────────────────────────────────────────────
    elif accion == "🗑️ Eliminar":
        if not data:
            st.markdown('<div class="info-box">No hay registros para eliminar.</div>', unsafe_allow_html=True)
            return

        opciones = {f"ID {r['id']} — {r.get(config['key'], '')}": r['id'] for r in data}
        sel = st.selectbox("Seleccionar registro a eliminar", list(opciones.keys()), key="sel_del")
        record_id = opciones[sel]

        st.markdown(f'<div class="error-box">⚠️ ¿Confirmas que deseas eliminar este registro? Esta acción no se puede deshacer.</div>', unsafe_allow_html=True)

        if st.button("ELIMINAR REGISTRO", key="btn_eliminar"):
            if delete_record(config['tabla'], record_id):
                st.markdown('<div class="success-box">✓ Registro eliminado</div>', unsafe_allow_html=True)
                st.rerun()
            else:
                st.markdown('<div class="error-box">Error al eliminar el registro</div>', unsafe_allow_html=True)
