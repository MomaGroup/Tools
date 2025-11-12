import streamlit as st

def aplicar_css_global():
    """
    Oculta el menú nativo de Streamlit y aplica estilos personalizados.
    Esta función debe llamarse DESPUÉS de crear el sidebar personalizado.
    """
    st.markdown("""
        <style>
        /* ============================================
           OCULTAR MENÚ NATIVO DE STREAMLIT
           ============================================ */
        
        /* Ocultar el panel de navegación completo */
        [data-testid="stSidebarNav"] {
            display: none !important;
            visibility: hidden !important;
            height: 0 !important;
            overflow: hidden !important;
        }
        
        /* Ocultar lista de páginas */
        [data-testid="stSidebarNav"] ul {
            display: none !important;
        }
        
        /* Ocultar el separador del menú */
        [data-testid="stSidebarNav"]::before,
        [data-testid="stSidebarNav"]::after {
            display: none !important;
        }
        
        /* Ocultar el botón de colapsar/expandir sidebar */
        [data-testid="collapsedControl"] {
            display: none !important;
        }
        
        /* ============================================
           ESTILOS DEL SIDEBAR PERSONALIZADO
           ============================================ */
        
        /* Ancho fijo del sidebar */
        section[data-testid="stSidebar"] {
            width: 320px !important;
            min-width: 320px !important;
        }
        
        /* Eliminar padding superior del sidebar */
        section[data-testid="stSidebar"] > div:first-child {
            padding-top: 1rem !important;
        }
        
        /* ============================================
           ESTILOS PARA BOTONES DEL MENÚ
           ============================================ */
        
        /* Botones del menú personalizado */
        div[data-testid="stSidebar"] button {
            border-radius: 8px !important;
            margin-bottom: 8px !important;
            font-weight: 500 !important;
            transition: all 0.3s ease !important;
        }
        
        /* Hover en botones */
        div[data-testid="stSidebar"] button:hover {
            transform: translateX(3px) !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;
        }
        
        /* ============================================
           OPCIONAL: OCULTAR ELEMENTOS ADICIONALES
           ============================================ */
        
        /* Ocultar header superior de Streamlit */
        header[data-testid="stHeader"] {
            display: none !important;
        }
        
        /* Ocultar el footer "Made with Streamlit" */
        footer {
            display: none !important;
        }
        
        /* Ocultar botón de menú hamburguesa en móvil */
        [data-testid="stToolbar"] {
            display: none !important;
        }
        
        /* ============================================
           ESTILOS PARA MEJORAR LA UI
           ============================================ */
        
        /* Espacio entre secciones */
        .block-container {
            padding-top: 2rem !important;
        }
        
        /* Estilo para los page_link */
        div[data-testid="stSidebar"] a {
            text-decoration: none !important;
            color: inherit !important;
        }
        
        /* Estilo para elementos anidados (submenús) */
        div[data-testid="stSidebar"] .element-container {
            padding-left: 0 !important;
        }
        </style>
    """, unsafe_allow_html=True)
