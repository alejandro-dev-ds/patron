import streamlit as st
from pathlib import Path

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

    if historico is None:
        st.error("Falta histórico de horas")
        st.stop()

    if proyectos is None:
        st.error("Falta tabla de proyectos")
        st.stop()

    if correcciones is None:
        st.error("Falta archivo de correcciones")
        st.stop()

    try:

        Path("historico.xlsx").write_bytes(
            historico.getbuffer()
        )

        Path("proyectos.xlsx").write_bytes(
            proyectos.getbuffer()
        )

        Path("correcciones.xlsx").write_bytes(
            correcciones.getbuffer()
        )

        with st.spinner("Generando patrón..."):

            patron, archivo_revision = ejecutar_patron(
                "historico.xlsx",
                "proyectos.xlsx",
                "correcciones.xlsx",
                "ejemplo_tabla_patron.xlsx"
            )

        st.success("Proceso finalizado correctamente")

        patron_bytes = Path(patron).read_bytes()

        st.download_button(
            label="📥 Descargar patrón",
            data=patron_bytes,
            file_name=Path(patron).name,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

        revision_bytes = Path(
            archivo_revision
        ).read_bytes()

        st.download_button(
            label="📥 Descargar correcciones actualizadas",
            data=revision_bytes,
            file_name=Path(archivo_revision).name,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    except Exception as e:
        st.error(f"Error durante la ejecución: {e}")
