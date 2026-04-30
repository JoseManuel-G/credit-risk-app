import streamlit as st

st.set_page_config(
    page_title="Análisis de Riesgo de Crédito",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Análisis de Riesgo de Crédito")

st.write(
    "Sube un Excel con Balance y PyG para calcular ratios financieros "
    "y generar un análisis de riesgo de crédito."
)

uploaded_file = st.file_uploader(
    "Sube tu archivo Excel",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    st.success("Archivo cargado correctamente.")

    st.write("Nombre del archivo:", uploaded_file.name)

    if st.button("Analizar empresa"):
        st.info("De momento la app funciona. En el siguiente paso conectaremos el motor financiero.")
