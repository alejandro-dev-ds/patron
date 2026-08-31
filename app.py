import streamlit as st

from generador_patron import ejecutar_patron

st.set_page_config(
    page_title="Generador de Patrón",
    layout="wide"
)

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

    st.success("Proceso finalizado")

    with open(patron, "rb") as f:
        st.download_button(
            "Descargar patrón",
            f,
            file_name=patron
        )

    with open(archivo_revision, "rb") as f:
        st.download_button(
            "Descargar correcciones actualizadas",
            f,
            file_name=archivo_revision
        )