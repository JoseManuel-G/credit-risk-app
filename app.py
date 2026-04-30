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
        with st.spinner("Analizando estados financieros..."):

            tmp_path = None

            try:
                # Guardamos temporalmente el Excel subido
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = tmp.name

                # Llamamos al motor financiero
                result = run_credit_risk_analysis(tmp_path)

                st.success("Análisis completado.")

                # =========================
                # RESUMEN EJECUTIVO
                # =========================
                st.header("Resumen ejecutivo")

                col1, col2, col3 = st.columns(3)

                with col1:
                    st.metric(
                        "Nivel de riesgo",
                        result.get("risk_level", "No calculado")
                    )

                with col2:
                    st.metric(
                        "Score",
                        result.get("risk_score", "No calculado")
                    )

                with col3:
                    st.metric(
                        "Calidad de datos",
                        result.get("data_quality", "No calculado")
                    )

                report_text = result.get("report_text")

                if report_text:
                    st.subheader("Análisis generado")
                    st.write(report_text)

                # =========================
                # RATIOS
                # =========================
                st.header("Ratios calculados")

                ratios = result.get("ratios")

                if isinstance(ratios, pd.DataFrame):
                    st.dataframe(ratios, use_container_width=True)
                elif isinstance(ratios, dict):
                    try:
                        st.dataframe(pd.DataFrame(ratios).T, use_container_width=True)
                    except Exception:
                        st.write(ratios)
                else:
                    st.info("Todavía no hay ratios calculados.")

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
                # DEBUG TÉCNICO
                # =========================
                with st.expander("Ver debug técnico"):
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
                st.error("Ha ocurrido un error durante el análisis.")
                st.exception(e)

            finally:
                if tmp_path and os.path.exists(tmp_path):
                    os.remove(tmp_path)
