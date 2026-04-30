import streamlit as st
import tempfile
import os
import pandas as pd

from motor_financiero import run_credit_risk_analysis


st.set_page_config(
    page_title="Análisis de Riesgo de Crédito",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Análisis de Riesgo de Crédito")

st.write(
    "Sube un Excel con Balance y PyG para generar un informe financiero "
    "de análisis de riesgo de crédito."
)

uploaded_file = st.file_uploader(
    "Sube tu archivo Excel",
    type=["xlsx", "xls"]
)

if uploaded_file is not None:
    st.success("Archivo cargado correctamente.")
    st.write("Nombre del archivo:", uploaded_file.name)

    if st.button("Generar informe"):
        with st.spinner("Generando informe financiero..."):

            tmp_path = None

            try:
                # Guardamos temporalmente el Excel subido
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name

                # Llamamos al motor financiero
                result = run_credit_risk_analysis(tmp_path)

                st.success("Informe generado correctamente.")

                # =========================
                # INFORME GENERADO
                # =========================
                st.header("Informe generado")

                report_text = result.get("report_text")

                if report_text:
                    st.write(report_text)
                else:
                    st.info("Todavía no se ha generado ningún informe.")

                # =========================
                # WARNINGS
                # =========================
                st.header("Warnings detectados")

                warnings = result.get("warnings", [])

                if warnings:
                    for warning in warnings:
                        st.warning(warning)
                else:
                    st.success("No se han detectado warnings relevantes.")

                # =========================
                # RATIOS / DATOS TÉCNICOS
                # =========================
                with st.expander("Ver ratios y datos técnicos"):
                    ratios = result.get("ratios")

                    st.subheader("Ratios calculados")

                    if isinstance(ratios, pd.DataFrame):
                        st.dataframe(ratios, use_container_width=True)
                    elif isinstance(ratios, dict):
                        try:
                            st.dataframe(pd.DataFrame(ratios).T, use_container_width=True)
                        except Exception:
                            st.write(ratios)
                    else:
                        st.info("Todavía no hay ratios calculados.")

                    st.subheader("Debug técnico")

                    ratios_debug_table = result.get("ratios_debug_table")

                    if isinstance(ratios_debug_table, pd.DataFrame):
                        st.dataframe(ratios_debug_table, use_container_width=True)
                    elif ratios_debug_table is not None:
                        st.write(ratios_debug_table)
                    else:
                        st.info("Todavía no hay debug técnico disponible.")

                # =========================
                # DESCARGAS
                # =========================
                st.header("Descargas")

                output_files = result.get("output_files", {})

                if output_files:
                    for file_label, file_path in output_files.items():
                        if file_path and os.path.exists(file_path):
                            with open(file_path, "rb") as f:
                                st.download_button(
                                    label=f"Descargar {file_label}",
                                    data=f,
                                    file_name=os.path.basename(file_path)
                                )
                else:
                    st.info("Todavía no hay archivos generados para descargar.")

            except Exception as e:
                st.error("Ha ocurrido un error durante la generación del informe.")
                st.exception(e)

            finally:
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
