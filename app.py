import streamlit as st
from generador_patron import ejecutar_patron

st.set_page_config(
    page_title="Generador de Patrón",
    layout="wide"
)

# Inicialización
if "patron_file" not in st.session_state:
    st.session_state.patron_file = None

if "revision_file" not in st.session_state:
    st.session_state.revision_file = None

if "patron_name" not in st.session_state:
    st.session_state.patron_name = None

if "revision_name" not in st.session_state:
    st.session_state.revision_name = None

st.title("Generador de patrón de horas")

st.write(
    "Sube los archivos necesarios y genera el patrón automáticamente."
)

historico = st.file_uploader(
    "Histórico de horas",
    type=["xlsx"]
)

proyectos = st.file_uploader(
    "Tabla de proyectos",
    type=["xlsx"]
)

correcciones = st.file_uploader(
    "Correcciones",
    type=["xlsx"]
)

if st.button("Generar patrón"):

    if not historico:
        st.error("Falta histórico de horas")
        st.stop()

    if not proyectos:
        st.error("Falta tabla de proyectos")
        st.stop()

    if not correcciones:
        st.error("Falta archivo de correcciones")
        st.stop()

    with open("historico.xlsx", "wb") as f:
        f.write(historico.getbuffer())

    with open("proyectos.xlsx", "wb") as f:
        f.write(proyectos.getbuffer())

    with open("correcciones.xlsx", "wb") as f:
        f.write(correcciones.getbuffer())

    patron, archivo_revision = ejecutar_patron(
        "historico.xlsx",
        "proyectos.xlsx",
        "correcciones.xlsx",
        "ejemplo_tabla_patron.xlsx"
    )

    # Guardar archivos en memoria
    with open(patron, "rb") as f:
        st.session_state.patron_file = f.read()

    with open(archivo_revision, "rb") as f:
        st.session_state.revision_file = f.read()

    st.session_state.patron_name = patron
    st.session_state.revision_name = archivo_revision

    st.success("Proceso finalizado")

# Mostrar siempre los botones si existen resultados
if st.session_state.patron_file is not None:

    st.download_button(
        label="📥 Descargar patrón",
        data=st.session_state.patron_file,
        file_name=st.session_state.patron_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if st.session_state.revision_file is not None:

    st.download_button(
        label="📥 Descargar correcciones actualizadas",
        data=st.session_state.revision_file,
        file_name=st.session_state.revision_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        on_click="ignore"
    )
