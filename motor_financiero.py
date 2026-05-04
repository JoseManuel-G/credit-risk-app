import pandas as pd


def run_credit_risk_analysis(file_path):
    """
    Primera prueba real del motor financiero.
    
    Objetivo:
    - Comprobar que Streamlit pasa bien el archivo.
    - Comprobar que el motor puede leer el Excel.
    - Devolver un informe básico en pantalla.
    """

    warnings = []

    try:
        # Leer nombres de hojas
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        # Intentar localizar Balance y PyG
        balance_sheet = None
        pyg_sheet = None

        for sheet in sheet_names:
            sheet_lower = sheet.lower()

            if "balance" in sheet_lower:
                balance_sheet = sheet

            if "pyg" in sheet_lower or "p&g" in sheet_lower or "perdidas" in sheet_lower or "ganancias" in sheet_lower:
                pyg_sheet = sheet

        if balance_sheet is None:
            warnings.append("No se ha detectado automáticamente la hoja de Balance.")

        if pyg_sheet is None:
            warnings.append("No se ha detectado automáticamente la hoja de PyG.")

        # Leer hojas si existen
        balance_shape = None
        pyg_shape = None

        if balance_sheet is not None:
            balance_df = pd.read_excel(file_path, sheet_name=balance_sheet)
            balance_shape = balance_df.shape

        if pyg_sheet is not None:
            pyg_df = pd.read_excel(file_path, sheet_name=pyg_sheet)
            pyg_shape = pyg_df.shape

        # Crear informe básico
        report_lines = []

        report_lines.append("## Informe financiero - prueba de conexión")
        report_lines.append("")
        report_lines.append("El archivo se ha recibido correctamente desde Streamlit y el motor financiero ha podido leerlo.")
        report_lines.append("")
        report_lines.append("### Hojas detectadas")
        report_lines.append(f"Hojas encontradas: {', '.join(sheet_names)}")
        report_lines.append("")

        if balance_sheet is not None:
            report_lines.append(f"Hoja de Balance detectada: **{balance_sheet}**")
            report_lines.append(f"Tamaño Balance: {balance_shape[0]} filas y {balance_shape[1]} columnas.")
        else:
            report_lines.append("No se ha podido identificar la hoja de Balance.")

        report_lines.append("")

        if pyg_sheet is not None:
            report_lines.append(f"Hoja de PyG detectada: **{pyg_sheet}**")
            report_lines.append(f"Tamaño PyG: {pyg_shape[0]} filas y {pyg_shape[1]} columnas.")
        else:
            report_lines.append("No se ha podido identificar la hoja de PyG.")

        report_lines.append("")
        report_lines.append("### Estado")
        report_lines.append(
            "Esta prueba confirma que el front ya está conectado con el motor. "
            "El siguiente paso será sustituir esta lectura básica por tu pipeline real: "
            "estructura de Excel, niveles, mapping, ratios e informe financiero."
        )

        report_text = "\n".join(report_lines)

        return {
            "warnings": warnings,
            "report_text": report_text,
            "ratios": None,
            "ratios_debug_table": None,
            "output_files": {}
        }

    except Exception as e:
        return {
            "warnings": ["Error al procesar el archivo."],
            "report_text": f"Ha ocurrido un error leyendo el Excel: {str(e)}",
            "ratios": None,
            "ratios_debug_table": None,
            "output_files": {}
        }
