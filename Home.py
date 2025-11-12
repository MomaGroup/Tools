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
            # ✅ OPCIÓN CON BOTÓN (más confiable)
            if st.button("🏦 Conciliación Bancaria", key="nav_conciliacion", use_container_width=True):
                try:
                    # Intenta con el nombre sin espacios primero
                    st.switch_page("pages/Conciliaciones/Conciliacion_bancaria.py")
                except:
                    try:
                        # Si falla, intenta con espacios
                        st.switch_page("pages/Conciliaciones/Conciliación bancaria.py")
                    except Exception as e:
                        st.error(f"⚠️ No se encuentra el archivo: {e}")

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
            # ✅ OPCIÓN CON BOTONES (más confiable que page_link)
            if st.button("📄 Formulario ICA Barranquilla", key="nav_ica", use_container_width=True):
                try:
                    st.switch_page("pages/Impuestos/1_Formulario_ICA_Barranquilla.py")
                except Exception as e:
                    st.error(f"⚠️ Error: {e}")
            
            if st.button("📄 Formulario Retefuente", key="nav_retefuente", use_container_width=True):
                try:
                    st.switch_page("pages/Impuestos/2_Formulario_Retefuente.py")
                except Exception as e:
                    st.error(f"⚠️ Error: {e}")
            
            if st.button("📄 Formulario SIMPLE", key="nav_simple", use_container_width=True):
                try:
                    st.switch_page("pages/Impuestos/3_Formulario_SIMPLE.py")
                except Exception as e:
                    st.error(f"⚠️ Error: {e}")

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
    
    # ==========================================================
    # 🔧 DIAGNÓSTICO DE ARCHIVOS (TEMPORAL - PARA DEBUG)
    # ==========================================================
    with st.expander("🔍 Diagnóstico del Sistema (Debug)"):
        import os
        st.write("📂 **Estructura de archivos detectada:**")
        
        try:
            if os.path.exists("pages"):
                for root, dirs, files in os.walk("pages"):
                    level = root.replace("pages", "").count(os.sep)
                    indent = " " * 2 * level
                    st.write(f"{indent}📁 {os.path.basename(root)}/")
                    sub_indent = " " * 2 * (level + 1)
                    for file in files:
                        if file.endswith('.py'):
                            ruta_completa = os.path.join(root, file)
                            st.write(f"{sub_indent}📄 {file}")
                            st.code(ruta_completa, language=None)
            else:
                st.error("⚠️ La carpeta 'pages' no existe en la raíz del proyecto")
                st.info("💡 Crea la carpeta 'pages' en la raíz del proyecto y coloca tus páginas allí")
        except Exception as e:
            st.error(f"Error al escanear archivos: {e}")

    st.markdown(f"""
    **Desarrollado por el área de Business Intelligence – Moma Group SAS**  

    *Sesión activa: {st.session_state.username}*
    """)

# ==========================================================
# 🐛 DEBUG: Mostrar información de session_state
# ==========================================================
# Descomenta estas líneas para ver el estado de la sesión durante desarrollo
# with st.expander("🐛 Debug - Session State"):
#     st.json(st.session_state.menu_state)
