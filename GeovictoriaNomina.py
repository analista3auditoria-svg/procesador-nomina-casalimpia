import streamlit as st
import pandas as pd
import re
import io

# Configuración de la página web amplia
st.set_page_config(
    page_title="Malla de Marcaciones GeoVictoria",
    page_icon="📊",
    layout="centered"
)

# Inicializar la memoria de sesión
if 'df_consolidado' not in st.session_state:
    st.session_state.df_consolidado = None
if 'df_vertical_final' not in st.session_state:
    st.session_state.df_vertical_final = None
if 'df_nomina_cargado' not in st.session_state:
    st.session_state.df_nomina_cargado = None
if 'periodos_disponibles' not in st.session_state:
    st.session_state.periodos_disponibles = []

# --- ESTILOS VISUALES PERSONALIZADOS (CSS) ---
st.markdown(
    """
    <style>
        .custom-header {
            background-color: #1e3a8a; 
            padding: 20px; 
            border-radius: 10px; 
            margin-bottom: 25px;
            text-align: center;
        }
        div[data-testid="stDataFrame"] table th {
            background-color: #1e3a8a !important;
            color: white !important;
            font-weight: bold !important;
            text-align: center !important;
            white-space: normal !important;
            vertical-align: middle !important;
            border: 1px solid #3b82f6 !important;
        }
        div[data-testid="stDataFrame"] {
            font-family: Arial, sans-serif;
        }
    </style>
    """, 
    unsafe_allow_html=True
)

# Banner corporativo superior
st.markdown(
    """
    <div class="custom-header">
        <h1 style="color:white; margin:0; font-family:Arial; font-size: 24px;">MALLA DE MARCACIONES GEOVICTORIA</h1>
        <p style="color:#cbd5e1; margin:5px 0 0 0; font-size: 14px;">Malla de Validación y Comparación de Horas - Casalimpia S.A.</p>
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

# --- PASO 1: CARGA DE ARCHIVOS ---
st.subheader("📁 1. Carga de Archivos Base")
col_file1, col_file2 = st.columns(2)

with col_file1:
    uploaded_file = st.file_uploader("Reporte Base GeoVictoria (.xlsx, .csv)", type=["xlsx", "csv"], key="gv_file")
with col_file2:
    uploaded_nomina = st.file_uploader("Archivo de Nómina (.xlsx, .csv)", type=["xlsx", "csv"], key="nom_file")

# Detectar y extraer periodos si el archivo de nómina es cargado
df_nom_preliminar = None
if uploaded_nomina is not None:
    try:
        if uploaded_nomina.name.endswith('.xlsx'):
            df_nom_preliminar = pd.read_excel(uploaded_nomina)
        else:
            try:
                df_nom_preliminar = pd.read_csv(uploaded_nomina, encoding='utf-8')
            except UnicodeDecodeError:
                df_nom_preliminar = pd.read_csv(uploaded_nomina, encoding='latin1')
        
        df_nom_preliminar.columns = [str(c).strip() for c in df_nom_preliminar.columns]
        
        # Buscar columna de periodo (Columna A o llamada PERIODO)
        col_periodo_name = None
        for c in df_nom_preliminar.columns:
            if c.upper() in ['PERIODO', 'MES', 'PERÍODO']:
                col_periodo_name = c
                break
        
        if col_periodo_name is None and len(df_nom_preliminar.columns) > 0:
            col_periodo_name = df_nom_preliminar.columns[0] # Por defecto columna A
            
        if col_periodo_name:
            st.session_state.periodos_disponibles = sorted(df_nom_preliminar[col_periodo_name].dropna().unique().tolist())
    except Exception:
        st.session_state.periodos_disponibles = []

# --- PASO 2: RANGO DE FECHAS Y FILTRO DE PERIODO ---
st.subheader("📅 2. Parámetros de Filtrado y Fechas")
col_f1, col_f2 = st.columns(2)
with col_f1:
    f_ini_input = st.date_input("Fecha Inicial Marcaciones (Desde)", pd.to_datetime("2026-01-10"))
with col_f2:
    f_fin_input = st.date_input("Fecha Final Marcaciones (Hasta)", pd.to_datetime("2026-02-13"))

# Desplegar el selector dinámico del periodo solicitado
periodos_seleccionados = st.multiselect(
    "📆 Seleccione el/los Periodos de Nómina a evaluar:",
    options=st.session_state.periodos_disponibles,
    default=st.session_state.periodos_disponibles,
    help="Filtra las filas del archivo de Nómina según los meses/periodos seleccionados."
)

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
                # --- LEER ARCHIVO GEOVICTORIA ---
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
                    st.error("❌ El archivo de GeoVictoria no contiene las columnas básicas necesarias (ID, Apellidos, Nombres, Fecha).")
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
                    st.error("❌ El archivo de GeoVictoria no contiene suficientes columnas (Mínimo hasta columna AW).")
                    st.stop()

                df['Fecha_Procesada'] = df['Fecha'].apply(parse_geovictoria_date)
                df = df.dropna(subset=['Fecha_Procesada'])
                df_filtered = df[(df['Fecha_Procesada'] >= f_ini) & (df['Fecha_Procesada'] <= f_fin)].copy()
                
                if df_filtered.empty:
                    st.warning("⚠️ No se encontraron registros de marcación en el rango de fechas seleccionado.")
                    st.stop()

                # --- CONTROL SEGURO PARA RFD, RFN, RDF, RNF (Evita el Error 'RFD') ---
                for fixed_col in ['RFD', 'RFN', 'RDF', 'RNF']:
                    if fixed_col not in df_filtered.columns:
                        df_filtered[fixed_col] = 0.0

                columnas_horas = ['COLUMNA_U', 'COLUMNA_W', 'COLUMNA_Y', 'COLUMNA_AA', 'COLUMNA_AC', 
                                  'COLUMNA_AE', 'COLUMNA_AG', 'COLUMNA_AI', 'COLUMNA_AK', 'COLUMNA_AM', 
                                  'COLUMNA_AO', 'COLUMNA_AQ', 'COLUMNA_AS', 'COLUMNA_AU', 'COLUMNA_AW',
                                  'RFD', 'RFN', 'RDF', 'RNF']
                                  
                for col in columnas_horas:
                    if col in df_filtered.columns:
                        df_filtered[col] = df_filtered[col].astype(str).str.replace(',', '.', regex=False).str.strip()
                        df_filtered[col] = df_filtered[col].replace(['', 'nan', 'None'], '0')
                        df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0.0)

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

                conceptos_tecnicos = ['C1067', 'C100730', 'C1066', 'C100729', 'C1060', 'C1061', 'C1063', 'C1062', 'C1064']
                df_grouped = df_filtered.groupby(['Identificador', 'Apellidos', 'Nombres'])[conceptos_tecnicos].sum().reset_index()

                columnas_multi = [
                    ('Número concepto', 'Identificador'), ('Número concepto', 'Apellidos'), ('Número concepto', 'Nombres'),
                    ('1067', 'Recargo Dominical No Compensado'), ('100730', 'Recargo Festivo'),
                    ('1066', 'Recargo Dominical Compensado'), ('100729', 'Recargo Festivo'),
                    ('1060', 'Recargo Nocturno 0.35%'), ('1061', 'Extras Ordinarias Diurnas'),
                    ('1063', 'Extras Festivas Diurnas'), ('1062', 'Extras Ordinarias Nocturnas'),
                    ('1064', 'Extras Festivas Nocturnas')
                ]
                df_grouped.columns = pd.MultiIndex.from_tuples(columnas_multi)
                st.session_state.df_consolidado = df_grouped
                
                # --- DESPIVOTE VERTICAL ---
                df_flat = df_grouped.copy()
                identificador_col = df_flat[('Número concepto', 'Identificador')]
                nombre_completo_col = df_flat[('Número concepto', 'Apellidos')].astype(str) + " " + df_flat[('Número concepto', 'Nombres')].astype(str)
                conceptos_columnas = [col for col in df_flat.columns if col[0] != 'Número concepto']
                
                df_listado = []
                for col in conceptos_columnas:
                    df_temp = pd.DataFrame({
                        'Identificador': identificador_col,
                        'Nombre': nombre_completo_col,
                        'Concepto': col[0].replace('C', ''),
                        'Nombre Concepto': col[1],
                        'Cantidad': df_flat[col]
                    })
                    df_listado.append(df_temp)
                    
                df_vertical_final = pd.concat(df_listado, ignore_index=True)
                df_vertical_final = df_vertical_final[df_vertical_final['Cantidad'] > 0].sort_values(by=['Identificador', 'Concepto']).reset_index(drop=True)
                st.session_state.df_vertical_final = df_vertical_final

                # --- PROCESAR NÓMINA CON FILTRO DE PERIODO ---
                if df_nom_preliminar is not None:
                    df_nom = df_nom_preliminar.copy()
                    
                    # Identificar la columna periodo dinámicamente
                    col_p = None
                    for c in df_nom.columns:
                        if c.upper() in ['PERIODO', 'MES', 'PERÍODO']: col_p = c; break
                    if col_p is None: col_p = df_nom.columns[0]
                    
                    # Aplicar filtro de periodos seleccionados por la interfaz
                    if periodos_seleccionados:
                        df_nom = df_nom[df_nom[col_p].isin(periodos_seleccionados)]
                    
                    rename_rules = {}
                    for c in df_nom.columns:
                        if c.lower() in ['identificador', 'id', 'cedula', 'cédula']: rename_rules[c] = 'Identificador'
                        if c.lower() in ['concepto', 'código', 'codigo']: rename_rules[c] = 'Concepto'
                        if c.lower() in ['cantidad', 'horas', 'valor_cantidad']: rename_rules[c] = 'Cantidad_Nomina'
                    
                    df_nom = df_nom.rename(columns=rename_rules)
                    df_nom = df_nom.rename(columns={col_p: 'PERIODO'})
                    
                    if 'Identificador' in df_nom.columns and 'Concepto' in df_nom.columns and 'Cantidad_Nomina' in df_nom.columns:
                        df_nom['Identificador'] = df_nom['Identificador'].astype(str).str.strip()
                        df_nom['Concepto'] = df_nom['Concepto'].astype(str).str.replace('C', '', regex=False).str.strip()
                        df_nom['Cantidad_Nomina'] = df_nom['Cantidad_Nomina'].astype(str).str.replace(',', '.', regex=False).str.strip()
                        df_nom['Cantidad_Nomina'] = pd.to_numeric(df_nom['Cantidad_Nomina'], errors='coerce').fillna(0.0)
                        df_nom['PERIODO'] = df_nom['PERIODO'].astype(str).str.strip()
                        
                        st.session_state.df_nomina_cargado = df_nom[['PERIODO', 'Identificador', 'Concepto', 'Cantidad_Nomina']]
                    else:
                        st.error("❌ El archivo de Nómina carece de la estructura requerida (Identificador, Concepto, Cantidad).")
                        st.session_state.df_nomina_cargado = None
                else:
                    st.session_state.df_nomina_cargado = None

                st.success("🎉 ¡Procesamiento completado con éxito!")
                
        except Exception as e:
            st.error(f"❌ Error en procesamiento: {str(e)}")

# --- SECCIÓN DE RESULTADOS ---
if st.session_state.df_consolidado is not None:
    st.markdown("---")
    
    # 1. MALLA
    st.subheader("🔍 1. Cuadro de Consulta General (Vista Malla)")
    busqueda = st.text_input("Filtrar por Identificador (Cédula) - Malla:", placeholder="Cédula...", key="s_malla").strip()
    df_m = st.session_state.df_consolidado
    if busqueda:
        df_m = df_m[df_m[('Número concepto', 'Identificador')].astype(str).str.contains(busqueda, case=False)]
    st.dataframe(df_m, use_container_width=True)
    
    # 2. DETALLE VERTICAL
    st.markdown("---")
    st.subheader("📋 2. Cuadro Detallado (Estructura Base GeoVictoria)")
    busqueda_v = st.text_input("Filtrar por Identificador (Cédula) - Detalle:", placeholder="Cédula...", key="s_vert").strip()
    df_v = st.session_state.df_vertical_final
    if busqueda_v:
        df_v = df_v[df_v['Identificador'].astype(str).str.contains(busqueda_v, case=False)]
    st.dataframe(df_v, use_container_width=True)

    # 3. COMPARATIVO CON PERIODO INCLUIDO
    st.markdown("---")
    st.subheader("⚖️ 3. Comparación Nómina vs Marcación")
    
    if st.session_state.df_nomina_cargado is not None:
        df_marcas = st.session_state.df_vertical_final.copy()
        df_marcas['Identificador'] = df_marcas['Identificador'].astype(str).str.strip()
        df_marcas['Concepto'] = df_marcas['Concepto'].astype(str).str.strip()
        
        df_nom_comp = st.session_state.df_nomina_cargado.copy()
        
        df_comparativo = pd.merge(
            df_marcas, 
            df_nom_comp, 
            on=['Identificador', 'Concepto'], 
            how='outer'
        )
        
        df_comparativo['Cantidad'] = df_comparativo['Cantidad'].fillna(0.0)
        df_comparativo['Cantidad_Nomina'] = df_comparativo['Cantidad_Nomina'].fillna(0.0)
        df_comparativo['Nombre'] = df_comparativo['Nombre'].fillna("No registrado en Marcaciones")
        df_comparativo['PERIODO'] = df_comparativo['PERIODO'].fillna("Sin Periodo")
        
        conceptos_dict = {"1060": "Recargo Nocturno 0.35%", "1066": "Recargo Dominical Compensado", 
                          "1067": "Recargo Dominical No Compensado", "100729": "Recargo Festivo", 
                          "100730": "Recargo Festivo", "1061": "Extras Ordinarias Diurnas", 
                          "1062": "Extras Ordinarias Nocturnas", "1063": "Extras Festivas Diurnas", 
                          "1064": "Extras Festivas Nocturnas"}
        
        df_comparativo['Nombre Concepto'] = df_comparativo['Nombre Concepto'].fillna(df_comparativo['Concepto'].map(conceptos_dict)).fillna("Concepto Desconocido")
        
        df_comparativo['DIF'] = df_comparativo['Cantidad_Nomina'] - df_comparativo['Cantidad']
        
        df_comparativo = df_comparativo.rename(columns={
            'PERIODO': 'PERIODO',
            'Cantidad_Nomina': 'Cantidad',
            'Cantidad': 'Geovictoria'
        })
        
        columnas_orden = ['PERIODO', 'Identificador', 'Nombre', 'Concepto', 'Nombre Concepto', 'Cantidad', 'Geovictoria', 'DIF']
        df_comparativo = df_comparativo[columnas_orden].sort_values(by=['PERIODO', 'Identificador', 'Concepto']).reset_index(drop=True)
        
        busqueda_c = st.text_input("Filtrar por Identificador (Cédula) - Comparativo:", placeholder="Cédula...", key="s_comp").strip()
        df_c_mostrar = df_comparativo
        if busqueda_c:
            df_c_mostrar = df_c_mostrar[df_c_mostrar['Identificador'].astype(str).str.contains(busqueda_c, case=False)]
            
        st.dataframe(df_c_mostrar, use_container_width=True)
        
        output_comp = io.BytesIO()
        with pd.ExcelWriter(output_comp, engine='openpyxl') as writer:
            df_comparativo.to_excel(writer, index=False, sheet_name="Diferencias_SIC")
        st.download_button(
            label="⚖️ DESCARGAR EXCEL DE COMPARACIÓN Y DIFERENCIAS",
            data=output_comp.getvalue(),
            file_name="Comparativo_Nomina_vs_Marcacion.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.warning("⚠️ Cargue el archivo de Nómina en el panel superior para activar la comparación.")
