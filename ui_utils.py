import streamlit as st

def aplicar_css_global():
    """
    CSS optimizado para alineación a la izquierda y espaciado compacto.
    """
    st.markdown("""
        <style>
        /* ============================================
           OCULTAR MENÚ NATIVO DE STREAMLIT
           ============================================ */
        
        [data-testid="stSidebarNav"] {
            display: none !important;
        }
        
        /* ============================================
           ESTILOS BÁSICOS DEL SIDEBAR
           ============================================ */
        
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #f8f9fa 0%, #ffffff 100%) !important;
        }
        
        section[data-testid="stSidebar"] > div:first-child {
            padding-top: 1rem !important;
            padding-left: 1rem !important;
            padding-right: 1rem !important;
        }
        
        /* ============================================
           ALINEACIÓN CRÍTICA - TODOS LOS ELEMENTOS
           ============================================ */
        
        /* Forzar TODO el sidebar a la izquierda */
        div[data-testid="stSidebar"],
        div[data-testid="stSidebar"] *,
        div[data-testid="stSidebar"] div,
        div[data-testid="stSidebar"] p,
        div[data-testid="stSidebar"] span {
            text-align: left !important;
            justify-content: flex-start !important;
            align-items: flex-start !important;
        }
        
        /* Contenedores de elementos */
        div[data-testid="stSidebar"] .element-container {
            text-align: left !important;
            display: block !important;
            margin-bottom: 0.15rem !important;
        }
        
        /* Contenedor de botones específicamente */
        div[data-testid="stSidebar"] .stButton {
            text-align: left !important;
            margin-bottom: 0.15rem !important;
        }
        
        div[data-testid="stSidebar"] .stButton > div {
            text-align: left !important;
            justify-content: flex-start !important;
        }
        
        /* ============================================
           ESTILOS PARA BOTONES - MUY ESPECÍFICO
           ============================================ */
        
        /* TODOS los botones del sidebar */
        div[data-testid="stSidebar"] button,
        div[data-testid="stSidebar"] button[kind="secondary"],
        div[data-testid="stSidebar"] button[kind="primary"] {
            width: 100% !important;
            text-align: left !important;
            justify-content: flex-start !important;
            display: flex !important;
            align-items: center !important;
            padding: 0.4rem 0.75rem !important;
            margin-bottom: 0.15rem !important;
            border-radius: 8px !important;
            font-weight: 500 !important;
            transition: all 0.3s ease !important;
            border: 1px solid transparent !important;
        }
        
        /* Contenido interno del botón */
        div[data-testid="stSidebar"] button > div,
        div[data-testid="stSidebar"] button > div > div,
        div[data-testid="stSidebar"] button p,
        div[data-testid="stSidebar"] button span {
            text-align: left !important;
            justify-content: flex-start !important;
            width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        
        /* Botón primario (Inicio) */
        div[data-testid="stSidebar"] button[kind="primary"] {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            color: white !important;
            font-weight: 600 !important;
        }
        
        /* Hover en botones */
        div[data-testid="stSidebar"] button:hover {
            transform: translateX(3px) !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;
            border-color: #e0e0e0 !important;
        }
        
        /* ============================================
           TÍTULOS Y TEXTO DEL SIDEBAR
           ============================================ */
        
        div[data-testid="stSidebar"] h1,
        div[data-testid="stSidebar"] h2,
        div[data-testid="stSidebar"] h3 {
            text-align: left !important;
            margin-top: 0.3rem !important;
            margin-bottom: 0.3rem !important;
            padding-left: 0 !important;
        }
        
        div[data-testid="stSidebar"] p,
        div[data-testid="stSidebar"] strong {
            text-align: left !important;
            margin-top: 0.3rem !important;
            margin-bottom: 0.1rem !important;
            padding-left: 0 !important;
            display: block !important;
        }
        
        /* Markdown específico */
        div[data-testid="stSidebar"] .stMarkdown {
            text-align: left !important;
            padding-left: 0 !important;
        }
        
        div[data-testid="stSidebar"] .stMarkdown > div {
            text-align: left !important;
        }
        
        /* ============================================
           SEPARADORES (HR)
           ============================================ */
        
        div[data-testid="stSidebar"] hr {
            margin: 0.8rem 0 !important;
            border: none !important;
            height: 1px !important;
            background: #e0e0e0 !important;
        }
        
        /* ============================================
           ESTILOS GENERALES DE LA APLICACIÓN
           ============================================ */
        
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }
        
        footer {
            display: none !important;
        }
        
        /* ============================================
           ESTILOS PARA ALERTAS Y MENSAJES
           ============================================ */
        
        .stSuccess {
            background-color: #d4edda !important;
            border-left: 4px solid #28a745 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
        
        .stWarning {
            background-color: #fff3cd !important;
            border-left: 4px solid #ffc107 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
        
        .stError {
            background-color: #f8d7da !important;
            border-left: 4px solid #dc3545 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
        
        .stInfo {
            background-color: #d1ecf1 !important;
            border-left: 4px solid #17a2b8 !important;
            border-radius: 8px !important;
            padding: 12px !important;
        }
        
        /* ============================================
           ESTILOS PARA TABLAS
           ============================================ */
        
        .dataframe {
            border-radius: 8px !important;
            overflow: hidden !important;
            border: 1px solid #e0e0e0 !important;
        }
        
        .dataframe thead tr {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
            color: white !important;
        }
        
        .dataframe thead th {
            font-weight: 600 !important;
            padding: 12px !important;
        }
        
        .dataframe tbody tr:nth-child(even) {
            background-color: #f8f9fa !important;
        }
        
        .dataframe tbody tr:hover {
            background-color: #e9ecef !important;
            transition: background-color 0.2s ease !important;
        }
        
        /* ============================================
           ESTILOS PARA INPUTS
           ============================================ */
        
        input, textarea, select {
            border-radius: 8px !important;
            border: 1px solid #ced4da !important;
            transition: all 0.3s ease !important;
        }
        
        input:focus, textarea:focus, select:focus {
            border-color: #667eea !important;
            box-shadow: 0 0 0 0.2rem rgba(102, 126, 234, 0.25) !important;
        }
        
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
           ESTILOS PARA BOTONES GENERALES (NO SIDEBAR)
           ============================================ */
        
        .stButton > button:not([data-testid="stSidebar"] button) {
            border-radius: 8px !important;
            font-weight: 500 !important;
            transition: all 0.3s ease !important;
            border: none !important;
            padding: 10px 24px !important;
        }
        
        .stButton > button:not([data-testid="stSidebar"] button):hover {
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
           RESPONSIVE
           ============================================ */
        
        @media (max-width: 768px) {
            .block-container {
                padding: 1rem !important;
            }
        }
        </style>
    """, unsafe_allow_html=True)


def aplicar_tema_oscuro():
    """
    Aplica un tema oscuro a la aplicación.
    """
    st.markdown("""
        <style>
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
        
        input, textarea, select {
            background-color: #2d2d2d !important;
            color: #e0e0e0 !important;
            border-color: #404040 !important;
        }
        
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
