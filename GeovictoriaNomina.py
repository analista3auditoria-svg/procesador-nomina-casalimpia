import streamlit as st
import pandas as pd
import re
import io

# Configuración de la página web
st.set_page_config(
    page_title="Procesador de Nómina - Casalimpia",
    page_icon="📊",
    layout="wide"
)

# Estilos visuales en la parte superior (Azul corporativo)
st.markdown(
    """
    <div style="background-color:#1e3a8a; padding:20px; border-radius:10px; margin-bottom:25px;">
        <h1 style="color:white; text-align:center; margin:0; font-family:Arial;">PROCESADOR DE NÓMINA - CASALIMPIA</h1>
        <p style="color:#cbd5e1; text-align:center; margin:5px 0 0 0;">Malla de Validación de Horas GeoVictoria</p>
    </div>
    """, 
    unsafe_allow_html=True  # <-- Corrección aquí (sin la 'ed')
)

# Funciones de procesamiento de fechas
def parse_geovictoria_date(val):
    if pd.isna(val):
        return pd.NaT
    match = re.search(r'(\d{2}-\d{2}-\d{4})', str(val))
    if match:
        return pd.to_datetime(match.group(1), format="%d-%m-%Y")
    return pd.NaT

# --- PASO 1: CARGA DE ARCHIVO ---
st.subheader("1. Seleccione el Reporte Base de GeoVictoria")
uploaded_file = st.file_uploader("Arrastra o selecciona tu archivo (.xlsx o .csv)", type=["xlsx", "csv"])

# --- PASO 2: RANGO DE FECHAS ---
st.subheader("2. Seleccione el Rango de Fechas")
col_f1, col_f2 = st.columns(2)
with col_f1:
    f_ini_input = st.date_input("Fecha Inicial (Desde)", pd.to_datetime("2026-01-10"))
with col_f2:
    f_fin_input = st.date_input("Fecha Final (Hasta)", pd.to_datetime("2026-02-13"))

st.info("Compilado optimizado para entorno Web Navegador. Los datos se procesan de forma segura.")

# --- PROCESAMIENTO ---
if uploaded_file is not None:
    if st.button("PROCESAR Y CUADRAR INFORMACIÓN", type="primary", use_container_width=True):
        try:
            f_ini = pd.to_datetime(f_ini_input)
            f_fin = pd.to_datetime(f_fin_input)
            
            if f_ini > f_fin:
                st.error("Error: La fecha inicial no puede ser mayor a la fecha final.")
            else:
                # Lectura según el tipo de archivo
                if uploaded_file.name.endswith('.xlsx'):
                    try:
                        df = pd.read_excel(uploaded_file, sheet_name="Reporte Geo victoria")
                    except ValueError:
                        st.error("No se encontró la pestaña 'Reporte Geo victoria' en el archivo Excel.")
                        st.stop()
                else:
                    try:
                        df = pd.read_csv(uploaded_file, encoding='utf-8')
                    except UnicodeDecodeError:
                        df = pd.read_csv(uploaded_file, encoding='latin1')

                # Homologación de cédula
                if 'ID' in df.columns:
                    df = df.rename(columns={'ID': 'Identificador'})
                    
                required = ['Identificador', 'Apellidos', 'Nombres', 'Fecha']
                if not all(col in df.columns for col in required):
                    st.error("El archivo cargado no contiene las columnas básicas necesarias (ID/Identificador, Apellidos, Nombres, Fecha).")
                    st.stop()

                # Mapeo de columnas por posición (Índices Base 0)
                if len(df.columns) >= 49:
                    df = df.rename(columns={
                        df.columns[20]: 'COLUMNA_U',   df.columns[22]: 'COLUMNA_W', 
                        df.columns[24]: 'COLUMNA_Y',   df.columns[26]: 'COLUMNA_AA',
                        df.columns[28]: 'COLUMNA_AC', df.columns[30]: 'COLUMNA_AE', 
                        df.columns[32]: 'COLUMNA_AG', df.columns[34]: 'COLUMNA_AI', 
                        df.columns[36]: 'COLUMNA_AK', df.columns[38]: 'COLUMNA_AM', 
                        df.columns[40]: 'COLUMNA_AO', df.columns[42]: 'COLUMNA_AQ', 
                        df.columns[44]: 'COLUMNA_AS', df.columns[46]: 'COLUMNA_AU', 
                        df.columns[48]: 'COLUMNA_AW'
                    })
                else:
                    st.error("El archivo no contiene suficientes columnas (Se requiere al menos hasta la columna AW).")
                    st.stop()

                # Filtrar por fechas
                df['Fecha_Procesada'] = df['Fecha'].apply(parse_geovictoria_date)
                df = df.dropna(subset=['Fecha_Procesada'])
                df_filtered = df[(df['Fecha_Procesada'] >= f_ini) & (df['Fecha_Procesada'] <= f_fin)].copy()
                
                if df_filtered.empty:
                    st.warning("No se encontraron registros en el rango de fechas seleccionado.")
                    st.stop()

                # Limpieza de datos numéricos
                columnas_horas = ['COLUMNA_U', 'COLUMNA_W', 'COLUMNA_Y', 'COLUMNA_AA', 'COLUMNA_AC', 
                                  'COLUMNA_AE', 'COLUMNA_AG', 'COLUMNA_AI', 'COLUMNA_AK', 'COLUMNA_AM', 
                                  'COLUMNA_AO', 'COLUMNA_AQ', 'COLUMNA_AS', 'COLUMNA_AU', 'COLUMNA_AW']
                                  
                for col in columnas_horas:
                    if col in df_filtered.columns:
                        df_filtered[col] = df_filtered[col].astype(str).str.replace(',', '.', regex=False).str.strip()
                        df_filtered[col] = df_filtered[col].replace(['', 'nan', 'None'], '0')
                        df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0.0)
                    else:
                        df_filtered[col] = 0.0

                for fixed_col in ['RFD', 'RFN', 'RDF', 'RNF']:
                    if fixed_col in df_filtered.columns:
                        df_filtered[fixed_col] = df_filtered[fixed_col].astype(str).str.replace(',', '.', regex=False).str.strip()
                        df_filtered[fixed_col] = pd.to_numeric(df_filtered[fixed_col], errors='coerce').fillna(0.0)
                    else:
                        df_filtered[fixed_col] = 0.0

                # Aplicación de fórmulas exactas de los recargos
                df_filtered['TOTAL DOM PLENO \n(1.75%)'] = df_filtered['COLUMNA_AM'] + df_filtered['COLUMNA_AO']
                df_filtered['TOTAL FEST (1.75%)'] = df_filtered['RFD'] + df_filtered['RFN'] + df_filtered['RDF'] + df_filtered['RNF']
                df_filtered['TOTAL DOM COMP \n(0.75%)'] = df_filtered['COLUMNA_AI'] + df_filtered['COLUMNA_AK']
                df_filtered['TOTAL FEST (0.75%)'] = (df_filtered['COLUMNA_AQ'] + df_filtered['COLUMNA_AS'] + 
                                                     df_filtered['COLUMNA_AU'] + df_filtered['COLUMNA_AW'])
                df_filtered['TOTAL REC.   NOC.    \n(0.35%)'] = (df_filtered['COLUMNA_AG'] + df_filtered['COLUMNA_AK'] + 
                                                                   df_filtered['COLUMNA_AS'] + df_filtered['COLUMNA_AO'] + 
                                                                   df_filtered['COLUMNA_AW'])
                df_filtered['EXTRAS ORDINARIAS DIURNAS (1.25%)'] = df_filtered['COLUMNA_U']
                df_filtered['EXTRAS DOMINICALES DIURNAS (2.00%)'] = df_filtered['COLUMNA_Y'] + df_filtered['COLUMNA_AC']
                df_filtered['EXTRAS ORDINARIAS NOCTURNAS (1.75%)'] = df_filtered['COLUMNA_W']
                df_filtered['EXTRAS DOMINICALES NOCTURNAS (2.50%)'] = df_filtered['COLUMNA_AA'] + df_filtered['COLUMNA_AE']

                conceptos_finales = [
                    'TOTAL DOM PLENO \n(1.75%)', 'TOTAL FEST (1.75%)', 'TOTAL DOM COMP \n(0.75%)',
                    'TOTAL FEST (0.75%)', 'TOTAL REC.   NOC.    \n(0.35%)', 'EXTRAS ORDINARIAS DIURNAS (1.25%)',
                    'EXTRAS DOMINICALES DIURNAS (2.00%)', 'EXTRAS ORDINARIAS NOCTURNAS (1.75%)', 'EXTRAS DOMINICALES NOCTURNAS (2.50%)'
                ]

                # Agrupación por empleado
                df_consolidado = df_filtered.groupby(['Identificador', 'Apellidos', 'Nombres'])[conceptos_finales].sum().reset_index()
                
                # Éxito: Mostrar vista previa en la página
                st.success("¡Cálculo completado con éxito!")
                st.dataframe(df_consolidado.head(10)) # Muestra los primeros 10 para verificar
                
                # Crear el archivo Excel en memoria para la descarga web
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_consolidado.to_excel(writer, index=False, sheet_name="Vista en SIC")
                processed_data = output.getvalue()
                
                # Botón web nativo de descarga
                st.download_button(
                    label="📥 DESCARGAR EXCEL CONSOLIDADO",
                    data=processed_data,
                    file_name="Consolidado_Nomina_SIC.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"Ocurrió un error inesperado: {str(e)}")
