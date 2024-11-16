import streamlit as st
from docx import Document
from docx.shared import Inches
from io import BytesIO
import os

# Ruta de la imagen de la firma (reemplaza con la ruta correcta en tu sistema)
firma_imagen_ruta = 'firma.jpg'

# Verificar si la imagen de firma existe
if not os.path.exists(firma_imagen_ruta):
    st.error("La imagen de la firma no se encuentra en el directorio.")
else:
    # Título de la página
    st.title("Generador de Documentos Word con Firma")

    # Crear una barra lateral para la navegación entre pantallas
    seleccion = st.sidebar.selectbox(
        "Selecciona una opción",
        ["Buscador ficha", "Generar Documento Word"]
    )

    if seleccion == "Buscador ficha":
        # Sección de inicio
        st.write("Bienvenido al generador de documentos Word con firma. Usa la barra lateral para seleccionar una opción.")

        # Descripción de la página para redirigir a Excel
        st.write("Ingrese su número de ficha para redirigir al archivo Excel asociado:")

        # Campo de entrada para el número de ficha
        ficha_id = st.text_input("Número de ficha:", placeholder="Ingrese el número de ficha")

        # Diccionario de IDs de ficha con sus URLs correspondientes
        links = {
            "2876293": "https://1drv.ms/x/c/7f0aa27d6eb5cc8f/EVm8_aNEU-1HklqlpaEjqEEBCdR2uXaVwdCzcZhgbPV9Pg?e=DDU7ZC",
            "2368753": "https://1drv.ms/x/c/7f0aa27d6eb5cc8f/Ee2vuJfHdZJLvRTDo-l0BhcBOnzGvKszPeCI1OTo_ta4wQ?e=6lOpEM",
            # Agrega más IDs de ficha y enlaces según sea necesario
        }

        # Botón para redirigir al enlace del Excel
        if st.button("Redirigir a su Excel de Ficha"):
            if ficha_id in links:
                st.success(f"Redirigiendo a Excel de la ficha {ficha_id}...")
                st.markdown(f"[Abrir Excel de Ficha {ficha_id}]({links[ficha_id]})", unsafe_allow_html=True)
            else:
                st.error("Ficha no encontrada. Por favor, verifica el número de ficha.")

    elif seleccion == "Generar Documento Word":
        # Sección para generar el documento Word
        st.write("Genera tu documento Word con la firma del instructor")

        # Campo de texto para ingresar el nombre
        nombre = st.text_input("Nombre:", placeholder="Ingrese su nombre")

        # Botón para generar y descargar el Word
        if st.button("Generar y Descargar Word"):
            if not nombre.strip():
                st.error("Por favor, ingresa un nombre.")
            else:
                # Crear un documento Word
                doc = Document()

                # Agregar texto al documento
                doc.add_paragraph(f"Este documento es del aprendiz: {nombre}.")

                # Insertar la imagen de la firma
                try:
                    doc.add_paragraph("Firma del instructor:")
                    doc.add_picture(firma_imagen_ruta, width=Inches(2))  # Ajusta el tamaño de la imagen si es necesario
                except Exception as e:
                    st.error(f"Error al insertar la firma: {str(e)}")

                # Agregar línea para la firma del instructor
                doc.add_paragraph("Firma del instructor: _______________")

                # Guardar el archivo Word en un buffer en memoria
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)

                # Proporcionar el archivo para descarga
                st.download_button(
                    label="Descargar Documento Word",
                    data=buffer,
                    file_name="firma_documento.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    else:
        st.error("Opción no válida seleccionada.")
