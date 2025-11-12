import streamlit as st
import hashlib

# ==========================================================
# 🔐 CONFIGURACIÓN DE USUARIOS
# ==========================================================
USUARIOS = {
    "admin": hashlib.sha256("admin123".encode()).hexdigest(),
    "usuario1": hashlib.sha256("password1".encode()).hexdigest(),
    # Agrega más usuarios aquí
}

# ==========================================================
# 🔐 FUNCIÓN DE LOGIN
# ==========================================================
def login():
    """
    Muestra formulario de login y verifica credenciales.
    Si el usuario ya está autenticado, no muestra nada.
    """
    # Inicializar estado de autenticación
    if 'autenticado' not in st.session_state:
        st.session_state.autenticado = False
    
    if 'username' not in st.session_state:
        st.session_state.username = None
    
    # Si ya está autenticado, no mostrar formulario
    if st.session_state.autenticado:
        return
    
    # Mostrar formulario de login
    st.title("🔐 Sistema de Autenticación")
    st.markdown("### Inicia sesión para acceder al sistema")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        with st.form("login_form"):
            username = st.text_input("👤 Usuario", placeholder="Ingresa tu usuario")
            password = st.text_input("🔒 Contraseña", type="password", placeholder="Ingresa tu contraseña")
            submit = st.form_submit_button("Iniciar Sesión", use_container_width=True, type="primary")
            
            if submit:
                # Verificar credenciales
                password_hash = hashlib.sha256(password.encode()).hexdigest()
                
                if username in USUARIOS and USUARIOS[username] == password_hash:
                    st.session_state.autenticado = True
                    st.session_state.username = username
                    st.success("✅ Inicio de sesión exitoso")
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos")
        
        # Información de usuarios de prueba (eliminar en producción)
        with st.expander("ℹ️ Credenciales de prueba"):
            st.info("""
            **Usuario:** admin  
            **Contraseña:** admin123
            
            **Usuario:** usuario1  
            **Contraseña:** password1
            """)
    
    # Detener ejecución si no está autenticado
    st.stop()

# ==========================================================
# 🚪 FUNCIÓN DE LOGOUT
# ==========================================================
def logout():
    """
    Cierra la sesión del usuario y recarga la aplicación.
    """
    st.session_state.autenticado = False
    st.session_state.username = None
    st.rerun()
