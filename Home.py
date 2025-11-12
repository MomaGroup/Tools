# ==========================================================
# 🏠 Home.py — Panel principal Moma Group Tools
# ==========================================================

import streamlit as st
from auth import login, logout
from ui_utils import aplicar_css_global

# ==========================================================
# 🌐 CONFIGURACIÓN INICIAL
# ==========================================================
st.set_page_config(
    page_title="Generador de Formularios Tributarios",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================================
# 🔐 LOGIN (AUTENTICACIÓN)
# ==========================================================
login()

# ==========================================================
# 🧭 INICIALIZAR ESTADO DEL MENÚ
# ==========================================================
if 'menu_state' not in st.session_state:
    st.session_state.menu_state = {
        'conciliaciones_expanded': False,
        'impuestos_expanded': False,
        'selected_page': 'Inicio'
    }

# ==========================================================
# 🧭 SIDEBAR PERSONALIZADO
# ==========================================================
with st.sidebar:
    st.markdown("# 📋 Menú")
    st.markdown("---")

    # ===== INICIO =====
    if st.button("🏠 Inicio", use_container_width=True, key="btn_inicio",
                 type="primary" if st.session_state.menu_state['selected_page'] == 'Inicio' else "secondary"):
        st.session_state.menu_state['selected_page'] = 'Inicio'
        st.session_state.menu_state['conciliaciones_expanded'] = False
        st.session_state.menu_state['impuestos_expanded'] = False
        st.rerun()

    # ===== CONCILIACIONES (DESPLEGABLE) =====
    conciliaciones_icon = "▲" if st.session_state.menu_state['conciliaciones_expanded'] else "▼"
    if st.button(f"🏦 Conciliaciones {conciliaciones_icon}",
                 use_container_width=True,
                 key="btn_conciliaciones",
                 type="primary" if st.session_state.menu_state['conciliaciones_expanded'] else "secondary"):
        st.session_state.menu_state['conciliaciones_expanded'] = not st.session_state.menu_state['conciliaciones_expanded']
        st.session_state.menu_state['impuestos_expanded'] = False
        st.rerun()

    # --- Submenú Conciliaciones ---
    if st.session_state.menu_state['conciliaciones_expanded']:
        col1, col2 = st.columns([0.1, 0.9])
        with col2:
            # ⚠️ IMPORTANTE: Ajusta el nombre exacto de tu archivo
            st.page_link("pages/Conciliaciones/Conciliacion_bancaria.py",
                         label="• Conciliación Bancaria", icon="🏦")

    # ===== IMPUESTOS (DESPLEGABLE) =====
    impuestos_icon = "▲" if st.session_state.menu_state['impuestos_expanded'] else "▼"
    if st.button(f"💰 Impuestos {impuestos_icon}",
                 use_container_width=True,
                 key="btn_impuestos",
                 type="primary" if st.session_state.menu_state['impuestos_expanded'] else "secondary"):
        st.session_state.menu_state['impuestos_expanded'] = not st.session_state.menu_state['impuestos_expanded']
        st.session_state.menu_state['conciliaciones_expanded'] = False
        st.rerun()

    # --- Submenú Impuestos ---
    if st.session_state.menu_state['impuestos_expanded']:
        col1, col2 = st.columns([0.1, 0.9])
        with col2:
            st.page_link("pages/Impuestos/1_Formulario_ICA_Barranquilla.py",
                         label="• Formulario ICA Barranquilla", icon="📄")
            st.page_link("pages/Impuestos/2_Formulario_Retefuente.py",
                         label="• Formulario Retefuente", icon="📄")
            st.page_link("pages/Impuestos/3_Formulario_SIMPLE.py",
                         label="• Formulario SIMPLE", icon="📄")

    # ===== INFORMACIÓN DEL USUARIO Y CERRAR SESIÓN =====
    st.markdown("---")
    st.success(f"👤 Usuario: **{st.session_state.username}**")
    if st.button("🚪 Cerrar Sesión", use_container_width=True, key="btn_logout"):
        logout()

# ==========================================================
# 🎨 APLICAR ESTILO GLOBAL (DESPUÉS DEL SIDEBAR)
# ==========================================================
aplicar_css_global()

# ==========================================================
# CONTENIDO PRINCIPAL SEGÚN LA SELECCIÓN
# ==========================================================
selected = st.session_state.menu_state['selected_page']

# ==========================================================
# 🏠 INICIO
# ==========================================================
if selected == 'Inicio':
    st.title("🏛️ Sistema de Generación de Formularios Tributarios")
    st.markdown("### Moma Group SAS")
    st.markdown("---")

    st.markdown("""
    ## Bienvenido al Sistema de Generación Automática de Formularios

    Este sistema te permite generar automáticamente formularios tributarios y conciliaciones 
    para múltiples empresas de forma rápida y eficiente.

    ### 📌 Secciones disponibles:

    - 🏦 **Conciliaciones**
    - 💰 **Formularios de Impuestos**

    Usa el menú lateral para seleccionar la funcionalidad que deseas utilizar.
    """)

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.info("""
        **✅ Formularios Tributarios**
        - ICA Barranquilla  
        - Retefuente  
        - Régimen SIMPLE  
        """)

    with col2:
        st.warning("""
        **🏦 Conciliaciones**
        - Conciliación Bancaria  
        - Cuentas por Cobrar (CxC)  
        - Cuentas por Pagar (CxP)  
        """)

    st.markdown("---")

    st.markdown("""
    ## 📖 Instrucciones de Uso

    1. Selecciona el formulario o conciliación desde el menú lateral  
    2. Sube los archivos requeridos (Excel o CSV según el módulo)  
    3. El sistema generará automáticamente el reporte o formulario en formato Excel  
    4. Descarga tus resultados con un solo clic

    ### ⚡ Características:
    - ✅ Procesamiento automático  
    - ✅ Descarga en formato Excel  
    - ✅ Interfaz moderna y segura  
    - ✅ Compatible con Streamlit Cloud  
    """)

    st.markdown("---")

    st.markdown(f"""
    **Desarrollado por el área de Business Intelligence – Moma Group SAS**  

    *Sesión activa: {st.session_state.username}*
    """)
