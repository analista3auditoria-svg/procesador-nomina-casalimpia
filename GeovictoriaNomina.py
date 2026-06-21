import streamlit as st
import pandas as pd
import re
import io

# Configuración de la página web estándar (centrada)
st.set_page_config(
    page_title="Procesador de Nómina - Casalimpia",
    page_icon="📊",
    layout="centered"
)

# Inicializar la memoria de sesión si no existe
if 'df_consolidado' not in st.session_state:
    st.session_state.df_consolidado = None

# --- CSS AVANZADO: DISEÑO CORPORATIVO PROFESIONAL ---
st.markdown(
    """
    <style>
        /* Fondo de la aplicación */
        .stApp {
            background-color: #f8fafc;
        }
        
        /* Banner superior estilizado */
        .custom-header {
            background: linear-gradient(135deg, #1e3a8a 0%, #0f172a 100%);
            padding: 30px; 
            border-radius: 12px; 
            margin-bottom: 30px;
            box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
        }
        
        /* Contenedores tipo tarjeta (Cards) */
        .step-card {
            background-color: #ffffff;
            padding: 24px;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            margin-bottom: 24px;
            box-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1);
        }
        
        /* Estilizar los encabezados de la tabla de Streamlit */
        div[data-testid="stDataFrame"] table th {
            background-color: #1e3a8a !important;
            color: white !important;
            font-weight: 600 !important;
            font-size: 13px !important;
            padding: 10px !important;
            text-align: center !important;
        }
        
        /* Rediseño de botones primarios */
        div.stButton > button[kind="primary"] {
            background: linear-gradient(90deg, #10b981 0%, #059669 100%) !important;
            border: none !important;
            padding: 12px 24px !important;
            font-weight: bold !important;
            transition: all 0.3s ease;
        }
        div.stButton > button[kind="primary"]:hover {
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(16, 185, 129, 0.3);
        }
    </style>
    """, 
    unsafe_allow_html=True
)

# Banner corporativo
st.markdown(
    """
    <div class="custom-header">
        <h1 style="color:white; text-align:center; margin:0; font-family:Arial; font-size:28px; letter-spacing:0.5px;">PROCESADOR DE NÓMINA - CASALIMPIA</h1>
        <p style="color:#94a3b8; text-align:center; margin:8px 0 0 0; font-size:14px;">Malla Avanzada de Validación de Horas GeoVictoria</p>
    </div>
    """, 
    unsafe_allow_html=True
)

def parse_geovictoria_date(val):
    if pd.isna(val):
        return pd.NaT
    match = re.search(r'(\d{2}-\d{2}-\d{4})', str(val))
    if match:
        return pd.to_datetime(match.group(1), format="%d-%m-%Y")
    return pd.NaT

# --- SECCIÓN DE ENTRADAS EN TARJETAS ---
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.subheader("📁 1. Seleccione el Reporte Base de GeoVictoria")
uploaded_file = st.file_uploader("Arrastra o selecciona tu archivo (.xlsx o .csv)", type=["xlsx", "csv"])
st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file is None:
    st.session_state.df_consolidado = None

st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.subheader("📅 2. Seleccione el Rango de Fechas")
col_f1, col_f2 = st.columns(2)
with col_f1:
    f_ini_input = st.date_input("Fecha Inicial (Desde)", pd.to_datetime("2026-01-10"))
with col_f2:
    f_fin_input = st.date_input("Fecha Final (Hasta)", pd.to_datetime("2026-02-13"))

st.info("Compilado optimizado para entorno Web Navegador. Los datos se procesan de forma segura.")
st.markdown('</div>', unsafe_allow_html=True)

# --- BOTÓN DE PROCESAMIENTO ---
if uploaded_file is not None:
    if st.button("PROCESAR Y CUADRAR INFORMACIÓN", type="primary", use_container_width=True):
        try:
            f_ini = pd.to_datetime(f_ini_input)
            f_fin = pd.to_datetime(f_fin_input)
            
            if f_ini > f_fin:
                st.error("Error: La fecha inicial no puede ser mayor a la fecha final.")
            else:
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

                if 'ID' in df.columns:
                    df = df.rename(columns={'ID': 'Identificador'})
                    
                required = ['Identificador', 'Apellidos', 'Nombres', 'Fecha']
                if not all(col in df.columns for col in required):
                    st.error("El archivo cargado no contiene las columnas básicas necesarias (ID/Identificador, Apellidos, Nombres, Fecha).")
                    st.stop()

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

                df['Fecha_Procesada'] = df['Fecha'].apply(parse_geovictoria_date)
                df = df.dropna(subset=['Fecha_Procesada'])
                df_filtered = df[(df['Fecha_Procesada'] >= f_ini) & (df['Fecha_Procesada'] <= f_fin)].copy()
                
                if df_filtered.empty:
                    st.warning("No se encontraron registros en el rango de fechas seleccionado.")
                    st.stop()

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

                # Fórmulas de recargos (Mantenemos los nombres exactos con sus saltos de línea \n originales)
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

                st.session_state.df_consolidado = df_filtered.groupby(['Identificador', 'Apellidos', 'Nombres'])[conceptos_finales].sum().reset_index()
                st.success("🎉 ¡Cálculo completado con éxito! Despliega hacia abajo para buscar y descargar.")
                
        except Exception as e:
            st.error(f"Ocurrió un error inesperado: {str(e)}")

# --- SECCIÓN DE RESULTADOS EN TARJETA ---
if st.session_state.df_consolidado is not None:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.subheader("🔍 Cuadro de Consulta y Búsqueda")
    
    busqueda = st.text_input(
        "Filtrar por Identificador (Cédula):", 
        placeholder="Escribe la cédula y presiona Enter para buscar..."
    ).strip()

    if busqueda:
        df_mostrar = st.session_state.df_consolidado[st.session_state.df_consolidado['Identificador'].astype(str).str.contains(busqueda, case=False)]
        if df_mostrar.empty:
            st.warning(f"No se encontró ningún empleado con la cédula: {busqueda}")
    else:
        df_mostrar = st.session_state.df_consolidado

    # Sistema de pestañas para organizar las tablas y que no desborden la pantalla
    tab1, tab2 = st.tabs(["📋 Vista Resumida", "📊 Detalle Completo de Recargos"])
    
    with tab1:
        # Remueve el índice numérico automático de Streamlit para ahorrar valioso espacio horizontal
        st.dataframe(df_mostrar[['Identificador', 'Apellidos', 'Nombres']], use_container_width=True, hide_index=True)
        
    with tab2:
        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
    
    # Preparar el archivo de descarga
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df_consolidado.to_excel(writer, index=False, sheet_name="Vista en SIC")
    processed_data = output.getvalue()
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        label="📥 DESCARGAR EXCEL CONSOLIDADO COMPLETO",
        data=processed_data,
        file_name="Consolidado_Nomina_SIC.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.markdown('</div>', unsafe_allow_html=True)
