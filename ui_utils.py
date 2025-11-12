import streamlit as st

def aplicar_css_global():
    """
    Oculta el menú nativo de Streamlit y aplica estilos personalizados.
    El sidebar permanece siempre visible y solo se contrae con el botón.
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
        
        /* ============================================
           SIDEBAR SIEMPRE VISIBLE Y FIJO - CORREGIDO
           ============================================ */
        
        /* Sidebar siempre expandido por defecto */
        section[data-testid="stSidebar"] {
            width: 320px !important;
            min-width: 320px !important;
            background: linear-gradient(180deg, #f8f9fa 0%, #ffffff 100%) !important;
            position: relative !important;
            transition: all 0.3s ease !important;
            overflow: visible !important;
        }
        
        /* Contenido del sidebar siempre visible cuando está expandido */
        section[data-testid="stSidebar"] > div {
            opacity: 1 !important;
            visibility: visible !important;
            transition: opacity 0.3s ease !important;
        }
        
        /* CORREGIDO: Cuando está colapsado (con el botón) */
        section[data-testid="stSidebar"][aria-expanded="false"] {
            width: 0px !important;
            min-width: 0px !important;
            margin-left: 0px !important;
            overflow: hidden !important;
        }
        
        /* Ocultar contenido del sidebar cuando está colapsado */
        section[data-testid="stSidebar"][aria-expanded="false"] > div {
            opacity: 0 !important;
            visibility: hidden !important;
            pointer-events: none !important;
            transition: opacity 0.2s ease !important;
        }
        
        /* Botón de colapsar/expandir sidebar */
        button[kind="header"] {
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
        }
        
        /* Botón de colapsar visible y estilizado */
        [data-testid="collapsedControl"] {
            display: flex !important;
            position: fixed !important;
            top: 0.5rem !important;
            left: 0.5rem !important;
            z-index: 999999 !important;
            background: #667eea !important;
            color: white !important;
            border-radius: 8px !important;
            padding: 8px !important;
            cursor: pointer !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15) !important;
            transition: all 0.2s ease !important;
            visibility: visible !important;
            opacity: 1 !important;
        }
        
        [data-testid="collapsedControl"]:hover {
            background: #5568d3 !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2) !important;
            transform: scale(1.05) !important;
        }
        
        /* Botón para colapsar cuando el sidebar está visible */
        [data-testid="stSidebar"] button[kind="header"] {
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
        }
        
        /* Eliminar padding superior del sidebar */
        section[data-testid="stSidebar"] > div:first-child {
            padding-top: 1rem !important;
        }
        
        /* Ajustar contenido cuando sidebar está expandido */
        .main .block-container {
            padding-left: 2rem !important;
            padding-right: 2rem !important;
            transition: all 0.3s ease !important;
        }
        
        /* NUEVO: Ajustar contenido cuando sidebar está colapsado */
        section[data-testid="stSidebar"][aria-expanded="false"] ~ .main .block-container {
            max-width: 100% !important;
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
            border: 1px solid transparent !important;
        }
        
        /* Hover en botones */
        div[data-testid="stSidebar"] button:hover {
            transform: translateX(3px) !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;
            border-color: #e0e0e0 !important;
        }
        
        /* Botones primarios activos */
        div[data-testid="stSidebar"] button[kind="primary"] {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            color: white !important;
            font-weight: 600 !important;
        }
        
        /* ============================================
           OPCIONAL: OCULTAR ELEMENTOS ADICIONALES
           ============================================ */
        
        /* Ocultar header superior de Streamlit (OPCIONAL) */
        /* Descomenta si quieres ocultar completamente el header */
        /* header[data-testid="stHeader"] {
            display: none !important;
        } */
        
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
            padding-bottom: 2rem !important;
        }
        
        /* Estilo para los page_link (si decides usarlos) */
        div[data-testid="stSidebar"] a {
            text-decoration: none !important;
            color: inherit !important;
            padding: 8px 12px !important;
            border-radius: 6px !important;
            display: block !important;
            transition: all 0.2s ease !important;
        }
        
        div[data-testid="stSidebar"] a:hover {
            background-color: #f0f2f6 !important;
            transform: translateX(3px) !important;
        }
        
        /* ============================================
           ESTILOS PARA ALERTAS Y MENSAJES
           ============================================ */
        
        /* Success messages */
        .stSuccess {
            background-color: #d4edda !important;
            border-left: 4px solid #28a745 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
        
        /* Warning messages */
        .stWarning {
            background-color: #fff3cd !important;
            border-left: 4px solid #ffc107 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
        
        /* Error messages */
        .stError {
            background-color: #f8d7da !important;
            border-left: 4px solid #dc3545 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
        
        /* Info messages */
        .stInfo {
            background-color: #d1ecf1 !important;
            border-left: 4px solid #17a2b8 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
        
        /* ============================================
           ESTILOS PARA TABLAS Y DATAFRAMES
           ============================================ */
        
        /* Tablas con bordes redondeados */
        .dataframe {
            border-radius: 8px !important;
            overflow: hidden !important;
            border: 1px solid #e0e0e0 !important;
        }
        
        /* Header de tablas */
        .dataframe thead tr {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            color: white !important;
        }
        
        .dataframe thead th {
            font-weight: 600 !important;
            padding: 12px !important;
        }
        
        /* Filas alternadas */
        .dataframe tbody tr:nth-child(even) {
            background-color: #f8f9fa !important;
        }
        
        .dataframe tbody tr:hover {
            background-color: #e9ecef !important;
            transition: background-color 0.2s ease !important;
        }
        
        /* ============================================
           ESTILOS PARA INPUTS Y FORMULARIOS
           ============================================ */
        
        /* Inputs con bordes redondeados */
        input, textarea, select {
            border-radius: 8px !important;
            border: 1px solid #ced4da !important;
            transition: all 0.3s ease !important;
        }
        
        input:focus, textarea:focus, select:focus {
            border-color: #667eea !important;
            box-shadow: 0 0 0 0.2rem rgba(102, 126, 234, 0.25) !important;
        }
        
        /* File uploader */
        [data-testid="stFileUploader"] {
            border: 2px dashed #ced4da !important;
            border-radius: 8px !important;
            padding: 20px !important;
            transition: all 0.3s ease !important;
        }
        
        [data-testid="stFileUploader"]:hover {
            border-color: #667eea !important;
            background-color: #f8f9fa !important;
        }
        
        /* ============================================
           ESTILOS PARA BOTONES GENERALES
           ============================================ */
        
        /* Botones principales */
        .stButton > button {
            border-radius: 8px !important;
            font-weight: 500 !important;
            transition: all 0.3s ease !important;
            border: none !important;
            padding: 10px 24px !important;
        }
        
        .stButton > button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15) !important;
        }
        
        /* ============================================
           ESTILOS PARA EXPANDERS
           ============================================ */
        
        .streamlit-expanderHeader {
            background-color: #f8f9fa !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
            transition: all 0.2s ease !important;
        }
        
        .streamlit-expanderHeader:hover {
            background-color: #e9ecef !important;
        }
        
        /* ============================================
           ESTILOS PARA SEPARADORES
           ============================================ */
        
        hr {
            margin: 2rem 0 !important;
            border: none !important;
            height: 2px !important;
            background: linear-gradient(90deg, transparent, #667eea, transparent) !important;
        }
        
        /* ============================================
           RESPONSIVE DESIGN
           ============================================ */
        
        @media (max-width: 768px) {
            section[data-testid="stSidebar"] {
                width: 280px !important;
                min-width: 280px !important;
            }
            
            .block-container {
                padding: 1rem !important;
            }
        }
        
        /* ============================================
           PREVENIR SCROLL HORIZONTAL
           ============================================ */
        
        body {
            overflow-x: hidden !important;
        }
        
        .main {
            overflow-x: hidden !important;
        }
        </style>
    """, unsafe_allow_html=True)


def aplicar_tema_oscuro():
    """
    Aplica un tema oscuro a la aplicación.
    Llamar esta función para cambiar al modo oscuro.
    """
    st.markdown("""
        <style>
        /* Modo Oscuro */
        .stApp {
            background-color: #1a1a1a !important;
            color: #e0e0e0 !important;
        }
        
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #2d2d2d 0%, #1a1a1a 100%) !important;
        }
        
        .block-container {
            background-color: #1a1a1a !important;
        }
        
        /* Inputs en modo oscuro */
        input, textarea, select {
            background-color: #2d2d2d !important;
            color: #e0e0e0 !important;
            border-color: #404040 !important;
        }
        
        /* Tablas en modo oscuro */
        .dataframe {
            background-color: #2d2d2d !important;
            color: #e0e0e0 !important;
        }
        
        .dataframe tbody tr:nth-child(even) {
            background-color: #262626 !important;
        }
        
        .dataframe tbody tr:hover {
            background-color: #333333 !important;
        }
        </style>
    """, unsafe_allow_html=True)
