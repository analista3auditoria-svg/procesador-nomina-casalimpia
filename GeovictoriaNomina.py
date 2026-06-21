import streamlit as st
import pandas as pd
import re
import io

# Configuración de la página web amplia
st.set_page_config(
    page_title="Malla de Marcaciones GeoVictoria",
    page_icon="📊",
    layout="wide"
)

# Inicializar la memoria de sesión si no existe
if 'df_consolidado' not in st.session_state:
    st.session_state.df_consolidado = None

# --- ESTILOS VISUALES PERSONALIZADOS (CSS) ---
st.markdown(
    """
    <style>
        /* Encabezado de la página centrado y limpio sin logo */
        .custom-header {
            background-color: #1e3a8a; 
            padding: 20px; 
            border-radius: 10px; 
            margin-bottom: 25px;
            text-align: center;
        }
        
        /* Personalización de las tablas de Streamlit (Soporte MultiIndex) */
        div[data-testid="stDataFrame"] table th {
            background-color: #1e3a8a !important;
            color: white !important;
            font-weight: bold !important;
            text-align: center !important;
            white-space: normal !important;
            vertical-align: middle !important;
            border: 1px solid #3b82f6 !important;
        }
        
        /* Ajuste para que los textos internos se alineen correctamente */
        div[data-testid="stDataFrame"] {
            font-family: Arial, sans-serif;
        }
    </style>
    """, 
    unsafe_allow_html=True
)

# Banner corporativo superior sin Logo
st.markdown(
    """
    <div class="custom-header">
        <h1 style="color:white; margin:0; font-family:Arial; font-size: 24px;">MALLA DE MARCACIONES GEOVICTORIA</h1>
        <p style="color:#cbd5e1; margin:5px 0 0 0; font-size: 14px;">Malla de Validación de Horas - Casalimpia S.A.</p>
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

# --- PASO 1: CARGA DE ARCHIVO ---
st.subheader("📁 1. Seleccione el Reporte Base de GeoVictoria")
uploaded_file = st.file_uploader("Arrastra o selecciona tu archivo (.xlsx o .csv)", type=["xlsx", "csv"])

if uploaded_file is None:
    st.session_state.df_consolidado = None

# --- PASO 2: RANGO DE FECHAS ---
st.subheader("📅 2. Seleccione el Rango de Fechas")
col_f1, col_f2 = st.columns(2)
with col_f1:
    f_ini_input = st.date_input("Fecha Inicial (Desde)", pd.to_datetime("2026-01-10"))
with col_f2:
    f_fin_input = st.date_input("Fecha Final (Hasta)", pd.to_datetime("2026-02-13"))

st.info("💡 Compilado optimizado para entorno Web Navegador. Los datos se procesan de forma segura.")

# --- BOTÓN DE PROCESAMIENTO ---
if uploaded_file is not None:
    if st.button("🚀 PROCESAR Y CUADRAR INFORMACIÓN", type="primary", use_container_width=True):
        try:
            f_ini = pd.to_datetime(f_ini_input)
            f_fin = pd.to_datetime(f_fin_input)
            
            if f_ini > f_fin:
                st.error("❌ Error: La fecha inicial no puede ser mayor a la fecha final.")
            else:
                if uploaded_file.name.endswith('.xlsx'):
                    try:
                        df = pd.read_excel(uploaded_file, sheet_name="Reporte Geo victoria")
                    except ValueError:
                        st.error("❌ No se encontró la pestaña 'Reporte Geo victoria' en el archivo Excel.")
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
                    st.error("❌ El archivo cargado no contiene las columnas básicas necesarias (ID/Identificador, Apellidos, Nombres, Fecha).")
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
                    st.error("❌ El archivo no contiene suficientes columnas (Se requiere al menos hasta la columna AW).")
                    st.stop()

                df['Fecha_Procesada'] = df['Fecha'].apply(parse_geovictoria_date)
                df = df.dropna(subset=['Fecha_Procesada'])
                df_filtered = df[(df['Fecha_Procesada'] >= f_ini) & (df['Fecha_Procesada'] <= f_fin)].copy()
                
                if df_filtered.empty:
                    st.warning("⚠️ No se encontraron registros en el rango de fechas seleccionado.")
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

                # --- OPERACIONES ARITMÉTICAS ---
                df_filtered['C1067'] = df_filtered['COLUMNA_AM'] + df_filtered['COLUMNA_AO']
                df_filtered['C100730'] = df_filtered['RFD'] + df_filtered['RFN'] + df_filtered['RDF'] + df_filtered['RNF']
                df_filtered['C1066'] = df_filtered['COLUMNA_AI'] + df_filtered['COLUMNA_AK']
                df_filtered['C100729'] = (df_filtered['COLUMNA_AQ'] + df_filtered['COLUMNA_AS'] + 
                                          df_filtered['COLUMNA_AU'] + df_filtered['COLUMNA_AW'])
                df_filtered['C1060'] = (df_filtered['COLUMNA_AG'] + df_filtered['COLUMNA_AK'] + 
                                        df_filtered['COLUMNA_AS'] + df_filtered['COLUMNA_AO'] + 
                                        df_filtered['COLUMNA_AW'])
                df_filtered['C1061'] = df_filtered['COLUMNA_U']
                df_filtered['C1063'] = df_filtered['COLUMNA_Y'] + df_filtered['COLUMNA_AC']
                df_filtered['C1062'] = df_filtered['COLUMNA_W']
                df_filtered['C1064'] = df_filtered['COLUMNA_AA'] + df_filtered['COLUMNA_AE']

                # Agrupación base inicial
                conceptos_tecnicos = ['C1067', 'C100730', 'C1066', 'C100729', 'C1060', 'C1061', 'C1063', 'C1062', 'C1064']
                df_grouped = df_filtered.groupby(['Identificador', 'Apellidos', 'Nombres'])[conceptos_tecnicos].sum().reset_index()

                # --- CONSTRUCCIÓN ESTRUCTURA DE DOBLE FILA (MultiIndex) ---
                columnas_multi = [
                    ('Número concepto', 'Identificador'),
                    ('Número concepto', 'Apellidos'),
                    ('Número concepto', 'Nombres'),
                    ('1067', 'TOTAL DOM PLENO (1.75%)'),
                    ('100730', 'TOTAL FEST (1.75%)'),
                    ('1066', 'TOTAL DOM COMP (0.75%)'),
                    ('100729', 'TOTAL FEST (0.75%)'),
                    ('1060', 'TOTAL REC. NOC. (0.35%)'),
                    ('1061', 'EXTRAS ORDINARIAS DIURNAS (1.25%)'),
                    ('1063', 'EXTRAS FESTIVAS DIURNAS (2.00%)'),
                    ('1062', 'EXTRAS ORDINARIAS NOCTURNAS (1.75%)'),
                    ('1064', 'EXTRAS FESTIVAS NOCTURNAS (2.50%)')
                ]
                
                # Asignamos el MultiIndex al dataframe finalizado
                df_grouped.columns = pd.MultiIndex.from_tuples(columnas_multi)
                st.session_state.df_consolidado = df_grouped
                
                st.success("🎉 ¡Cálculo completado con éxito! Despliega hacia abajo para buscar y descargar.")
                
        except Exception as e:
            st.error(f"❌ Ocurrió un error inesperado: {str(e)}")

# --- SECCIÓN DE RESULTADOS PERSISTENTE ---
if st.session_state.df_consolidado is not None:
    st.markdown("---")
    st.subheader("🔍 Cuadro de Consulta y Búsqueda")
    
    busqueda = st.text_input(
        "Filtrar por Identificador (Cédula):", 
        placeholder="Escribe la cédula y presiona Enter para buscar..."
    ).strip()

    if busqueda:
        df_mostrar = st.session_state.df_consolidado[st.session_state.df_consolidado[('Número concepto', 'Identificador')].astype(str).str.contains(busqueda, case=False)]
        if df_mostrar.empty:
            st.warning(f"⚠️ No se encontró ningún empleado con la cédula: {busqueda}")
    else:
        df_mostrar = st.session_state.df_consolidado

    # Tabla interactiva en la Web (Muestra perfectamente la doble fila)
    st.dataframe(df_mostrar, use_container_width=True)
    
    # --- PREPARACIÓN DEL EXCEL CORREGIDA (Evita NotImplementedError) ---
    df_excel = st.session_state.df_consolidado.copy()
    
    # Convertimos los nombres combinados a una sola fila limpia para el archivo descargable
    df_excel.columns = [f"{col[0]} - {col[1]}" for col in df_excel.columns]
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_excel.to_excel(writer, index=False, sheet_name="Vista en SIC")
    processed_data = output.getvalue()
    
    st.download_button(
        label="📥 DESCARGAR EXCEL CONSOLIDADO COMPLETO",
        data=processed_data,
        file_name="Consolidado_Nomina_SIC.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
