import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="Gestión de Rotables", layout="wide", page_icon="✈️")

# ==========================================
# SISTEMA DE AUTENTICACIÓN GLOBAL
# ==========================================
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("Acceso al Sistema")
        st.write("Por favor, ingresa la contraseña para continuar.")
        pwd = st.text_input("Contraseña:", type="password")
        
        if pwd == st.secrets.get("admin_password", ""):
            st.session_state["password_correct"] = True
            st.rerun()
        elif pwd != "":
            st.error("Contraseña incorrecta.")
        return False
    return True

# ==========================================
# APLICACIÓN PRINCIPAL
# ==========================================
if check_password():
    banner_html = """
    <style>
        .hero-banner {
            background: linear-gradient(135deg, #38E54D 0%, #E607E9 100%);
            padding: 2.5rem 2rem;
            border-radius: 12px;
            margin-bottom: 2rem;
            box-shadow: 0 6px 10px rgba(0, 0, 0, 0.15);
            color: white;
            text-align: left;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }
        .hero-title {
            font-size: 2.4rem;
            font-weight: 800;
            margin: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            letter-spacing: 0.5px;
        }
        .hero-subtitle {
            font-size: 1.1rem;
            font-weight: 400;
            margin-top: 8px;
            opacity: 0.9;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
    </style>

    <div class="hero-banner">
        <div class="hero-title">SISTEMA DE REPORTES DE ROTABLES</div>
    </div>
    """

    st.markdown(banner_html, unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["Captura de Datos", "Cruce de Información"])

    # --- PESTAÑA 1: CAPTURA DE DATOS ---
    with tab1:
        st.header("Captura de Reporte Manual")
        st.markdown("Ingresa la información de los componentes. Puedes agregar nuevas filas al final de la tabla.")

        columnas_plantilla = [
            "No.", "FECHA", "DESCRIPCIÓN", "NUMERO DE PARTE", "NUMERO DE SERIE", 
            "NUMERO DE PARTE2", "NUMERO DE SERIE3", "EST", "AC", "DOCS",
            "N°TO", "RO/EX", "REGIMEN", "OBSERVACIONES"
        ]
        
        if "df_input" not in st.session_state:
            df_vacio = pd.DataFrame(columns=columnas_plantilla)
            for i in range(5):
                df_vacio.loc[i] = [None] * 14
            
            df_vacio["FECHA"] = pd.to_datetime(df_vacio["FECHA"])
            
            for col in columnas_plantilla:
                if col != "FECHA":
                    df_vacio[col] = df_vacio[col].astype(str)

            st.session_state["df_input"] = df_vacio

        edited_df = st.data_editor(
            st.session_state["df_input"],
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "FECHA": st.column_config.DateColumn("FECHA", format="DD/MM/YYYY"),
                "DESCRIPCIÓN": st.column_config.TextColumn("DESCRIPCIÓN"),
                "NUMERO DE PARTE": st.column_config.TextColumn("INSTALADO: NO. DE PARTE"),
                "NUMERO DE SERIE": st.column_config.TextColumn("INSTALADO: NO. DE SERIE"),
                "NUMERO DE PARTE2": st.column_config.TextColumn("REMOVIDO: NO. DE PARTE"),
                "NUMERO DE SERIE3": st.column_config.TextColumn("REMOVIDO: NO. DE SERIE")
            }
        )

        st.markdown("---")
        
        output_captura = BytesIO()
        with pd.ExcelWriter(output_captura, engine='openpyxl') as writer:
            edited_df.to_excel(writer, index=False, sheet_name='Reporte_Manual')
        output_captura.seek(0)

        st.download_button(
            label="Descargar Reporte para Enviar",
            data=output_captura,
            file_name="Reporte_Manual_Capturado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    # --- PESTAÑA 2: CRUCE DE INFORMACIÓN ---
    with tab2:
        st.header("Cruce de Información")
        st.markdown("Sube el reporte manual capturado y el archivo en crudo descargado del sistema.")
        
        st.markdown("""
            <style>
            [data-testid="column"]:nth-of-type(1) div[data-testid="stFileUploadDropzone"] {
                background-color: #072a40 !important;
                border: 2px dashed #17a2b8 !important;
                border-radius: 10px !important;
            }
            [data-testid="column"]:nth-of-type(2) div[data-testid="stFileUploadDropzone"] {
                background-color: #053015 !important;
                border: 2px dashed #28a745 !important;
                border-radius: 10px !important;
            }
            </style>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            file_plantilla = st.file_uploader("1. Sube el Reporte Manual", type=["xlsx"])
        with col2:
            file_descargado = st.file_uploader("2. Sube el Archivo del Sistema", type=["xlsx"])

        if file_plantilla and file_descargado:
            st.info("Procesando cruce de datos...")
            try:
                df_plantilla = pd.read_excel(file_plantilla, dtype=str)
                df_descargado = pd.read_excel(file_descargado, dtype=str)
                
                df_plantilla.columns = df_plantilla.columns.str.strip()
                df_descargado.columns = df_descargado.columns.str.strip()

                # Compatibilidad con archivos históricos
                df_plantilla = df_plantilla.rename(columns={
                    "NUMERO DE PARTE.1": "NUMERO DE PARTE2",
                    "NUMERO DE SERIE.1": "NUMERO DE SERIE3"
                })

                if "NUMERO DE SERIE3" in df_plantilla.columns and "SN" in df_descargado.columns:
                    df_plantilla["NUMERO DE SERIE3"] = df_plantilla["NUMERO DE SERIE3"].astype(str).str.strip()
                    df_descargado["SN"] = df_descargado["SN"].astype(str).str.strip()

                    df_cruce = pd.merge(
                        df_plantilla,
                        df_descargado,
                        left_on="NUMERO DE SERIE3",
                        right_on="SN",
                        how="left"
                    )

                    if "PN_DESCRIP" in df_cruce.columns and "DESCRIPCIÓN" in df_cruce.columns:
                        df_cruce["DESCRIPCIÓN"] = np.where(
                            df_cruce["PN_DESCRIP"].notna() & (df_cruce["PN_DESCRIP"] != "nan") & (df_cruce["PN_DESCRIP"].str.strip() != ""), 
                            df_cruce["PN_DESCRIP"], 
                            df_cruce["DESCRIPCIÓN"]
                        )
                    
                    if "PN" in df_cruce.columns and "NUMERO DE PARTE2" in df_cruce.columns:
                        df_cruce["NUMERO DE PARTE2"] = np.where(
                            df_cruce["PN"].notna() & (df_cruce["PN"] != "nan") & (df_cruce["PN"].str.strip() != ""), 
                            df_cruce["PN"], 
                            df_cruce["NUMERO DE PARTE2"]
                        )

                    columnas_originales = df_plantilla.columns.tolist()
                    df_final = df_cruce[columnas_originales]
                    
                    df_final = df_final.replace(['nan', 'NaN', '#REF!', 'None', 'NaT'], '')
                    df_final = df_final.fillna('')

                    # Truco visual para el Excel de salida (espacio final)
                    df_final = df_final.rename(columns={
                        "NUMERO DE PARTE2": "NUMERO DE PARTE ", 
                        "NUMERO DE SERIE3": "NUMERO DE SERIE "
                    })

                    st.success("¡Cruce realizado con éxito!")
                    st.dataframe(df_final.head(5))

                    nivel_superior = [
                        "DATOS GENERALES", 
                        "DATOS GENERALES", "DATOS GENERALES", 
                        "COMPONENTE INSTALADO", "COMPONENTE INSTALADO",
                        "COMPONENTE REMOVIDO", "COMPONENTE REMOVIDO",  
                        "DATOS GENERALES", "DATOS GENERALES", "DATOS GENERALES", 
                        "DATOS GENERALES", "DATOS GENERALES", "DATOS GENERALES", "DATOS GENERALES" 
                    ]
                    
                    output_final = BytesIO()
                    with pd.ExcelWriter(output_final, engine='openpyxl') as writer:
                        df_final.to_excel(writer, index=False, sheet_name='Reporte_Completado', startrow=3)
                        
                        workbook = writer.book
                        worksheet = writer.sheets['Reporte_Completado']
                        
                        worksheet['A1'] = "REPORTE DE COMPONENTES ROTABLES"
                        worksheet['A1'].font = Font(size=16, bold=True, color="FFFFFF")
                        worksheet['A1'].alignment = Alignment(horizontal="center", vertical="center")
                        worksheet['A1'].fill = openpyxl.styles.PatternFill(start_color="002060", end_color="002060", fill_type="solid")
                        
                        ultima_col_letra = get_column_letter(len(df_final.columns))
                        worksheet.merge_cells(f'A1:{ultima_col_letra}1')
                        
                        for col_num, valor in enumerate(nivel_superior, 1):
                            celda = worksheet.cell(row=3, column=col_num)
                            celda.value = valor
                            celda.font = Font(bold=True)
                            celda.alignment = Alignment(horizontal="center", vertical="center")
                            celda.fill = openpyxl.styles.PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                        
                        ultima_fila = len(df_final) + 4 
                        rango_tabla = f"A4:{ultima_col_letra}{ultima_fila}"
                        
                        tabla = Table(displayName="TablaRotables", ref=rango_tabla)
                        estilo = TableStyleInfo(
                            name="TableStyleMedium9", 
                            showFirstColumn=False,
                            showLastColumn=False, 
                            showRowStripes=True, 
                            showColumnStripes=False
                        )
                        tabla.tableStyleInfo = estilo
                        worksheet.add_table(tabla)
                        
                        for col_idx in range(1, worksheet.max_column + 1):
                            column_letter = get_column_letter(col_idx)
                            max_length = 0
                            for row_idx in range(3, worksheet.max_row + 1):
                                cell = worksheet.cell(row=row_idx, column=col_idx)
                                try:
                                    if cell.value:
                                        longitud = len(str(cell.value))
                                        if longitud > max_length:
                                            max_length = longitud
                                except:
                                    pass
                            worksheet.column_dimensions[column_letter].width = max(max_length + 2, 10)

                    output_final.seek(0)

                    st.download_button(
                        label="Descargar Reporte Finalizado",
                        data=output_final,
                        file_name="Reporte_Rotables_Completado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                else:
                    st.error("No se encontraron las columnas necesarias para el cruce. Verifica los archivos subidos.")
            except Exception as e:
                st.error(f"Error procesando los datos: {e}")
