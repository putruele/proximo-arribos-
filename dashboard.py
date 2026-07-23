import os
import io
import json
import requests
import pandas as pd
import streamlit as st
from datetime import datetime, date

# Configuración de página de Streamlit
st.set_page_config(
    page_title="Dashboard de Próximos Arribos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para una estética premium
st.markdown("""
<style>
    /* Estilos generales y limitación de ancho para pantallas de laptops */
    .main {
        background-color: #f8f9fa;
    }
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
        max-width: 1450px !important;
        margin: 0 auto;
    }
    h1 {
        color: #1e293b;
        font-family: 'Outfit', 'Inter', sans-serif;
        font-weight: 700;
        font-size: 26px !important;
        margin-bottom: 8px !important;
    }
    h3 {
        font-size: 18px !important;
        margin-top: 8px !important;
        margin-bottom: 8px !important;
    }
    /* Tarjetas KPI compactas para evitar comer espacio vertical en notebooks */
    .metric-card {
        background-color: white;
        border-radius: 10px;
        padding: 12px 15px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
        border: 1px solid #e2e8f0;
        text-align: center;
    }
    .metric-title {
        font-size: 11px;
        color: #64748b;
        font-weight: 600;
        text-transform: uppercase;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 22px;
        color: #0f172a;
        font-weight: 700;
    }
    /* Reducir tamaño de texto en la grilla st.dataframe para que calce en 15 pulgadas */
    div[data-testid="stDataFrame"] {
        font-size: 13px !important;
    }
</style>
""", unsafe_allow_html=True)

# Archivo de datos consolidado
DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "datos_consolidados.csv")

def parse_numeric_value(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace('$', '').replace(' ', '')
    if not val_str:
        return 0.0
    if ',' in val_str and '.' in val_str:
        if val_str.find('.') < val_str.find(','):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

# Función para cargar datos (soporta lectura en la nube de Google Sheets con fallback local)
@st.cache_data(ttl=300)
def load_data(gsheet_url, file_mtime):
    df = None
    if gsheet_url:
        try:
            # Forzar la carga mediante pandas leyendo la URL del export del sheet
            df = pd.read_csv(gsheet_url)
        except Exception as e:
            st.warning(f"No se pudo cargar desde Google Sheets ({e}). Cargando archivo local...")
            
    # Fallback local
    if df is None and os.path.exists(DATA_PATH):
        try:
            df = pd.read_csv(DATA_PATH)
        except Exception as e:
            st.error(f"Error al leer consolidado local: {e}")
            
    if df is not None:
        try:
            df['fecha'] = pd.to_datetime(df['fecha'])
            numeric_cols = [
                'kg_unitario', 'cantidad', 'costo_final_unitario', 'costo_por_kilo',
                'kilos_totales', 'costo_final_total'
            ]
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].apply(parse_numeric_value)
            return df
        except Exception as e:
            st.error(f"Error al procesar tipos de datos: {e}")
            
    return None

# Función para persistir nuevas filas (soporta Google Sheets via Apps Script Web App o localmente en el CSV)
def save_new_rows(new_rows_df):
    apps_script_url = st.secrets.get("APPS_SCRIPT_URL", None)
    db_cols = [
        'fecha', 'nro_fc', 'guia', 'articulo_original', 'kg_unitario', 'cantidad',
        'costo_final_unitario', 'costo_por_kilo', 'proveedor', 'archivo_origen',
        'articulo_producto', 'articulo_variedad', 'kilos_totales', 'costo_final_total'
    ]
    new_rows_df = new_rows_df[db_cols].copy()
    new_rows_df['fecha'] = new_rows_df['fecha'].dt.strftime('%Y-%m-%d')
    
    if apps_script_url:
        try:
            rows_to_send = new_rows_df.values.tolist()
            response = requests.post(apps_script_url, data=json.dumps(rows_to_send))
            if response.status_code == 200 and response.text == "Success":
                return True
            else:
                st.error(f"Error del Apps Script: {response.text}")
                return False
        except Exception as e:
            st.error(f"Error al conectar con Google Sheets: {e}")
            return False
    else:
        # Fallback local
        try:
            if os.path.exists(DATA_PATH):
                existing_df = pd.read_csv(DATA_PATH)
                updated_df = pd.concat([existing_df, new_rows_df], ignore_index=True)
            else:
                updated_df = new_rows_df
            updated_df.to_csv(DATA_PATH, index=False, encoding='utf-8-sig')
            return True
        except Exception as e:
            st.error(f"Error al guardar archivo consolidado local: {e}")
            return False

# Funciones auxiliares para consolidación en memoria
def clean_header_name(h):
    if not h:
        return ""
    return " ".join(str(h).split()).upper().strip()

def map_and_clean_df(df):
    if df.empty:
        return pd.DataFrame()
        
    mapped_df = pd.DataFrame()
    cols = df.columns
    
    col_fecha = [c for c in cols if 'FECHA' in c or 'DATE' in c]
    if col_fecha:
        mapped_df['fecha'] = df[col_fecha[0]]
    else:
        mapped_df['fecha'] = pd.Series([None]*len(df))
        
    import re
    col_fc = []
    for c in cols:
        if re.search(r'\bFC\b|\bFC\W|FACTURA|INVOICE', c):
            col_fc.append(c)
            break
    if col_fc:
        mapped_df['nro_fc'] = df[col_fc[0]]
    else:
        mapped_df['nro_fc'] = pd.Series([None]*len(df))
        
    col_guia = [c for c in cols if 'GUIA' in c or 'LOTE' in c or 'TRACKING' in c]
    if col_guia:
        mapped_df['guia'] = df[col_guia[0]]
    else:
        mapped_df['guia'] = pd.Series([None]*len(df))
        
    col_art = [c for c in cols if 'ARTICULO' in c or 'PRODUCTO' in c or 'PRODUCT' in c or 'ITEM' in c]
    if col_art:
        mapped_df['articulo_original'] = df[col_art[0]]
    else:
        mapped_df['articulo_original'] = pd.Series([None]*len(df))
        
    col_kg = []
    for c in cols:
        if ('KG' in c or 'KILO' in c) and not any(k in c for k in ['COSTO', 'PRECIO', 'PESOS', 'PAGAR']):
            col_kg.append(c)
            break
    if col_kg:
        mapped_df['kg_unitario'] = df[col_kg[0]]
    else:
        mapped_df['kg_unitario'] = pd.Series([None]*len(df))
        
    col_cant = [c for c in cols if 'CANT' in c or 'QTY' in c or 'CANTIDAD' in c]
    if col_cant:
        mapped_df['cantidad'] = df[col_cant[0]]
    else:
        mapped_df['cantidad'] = pd.Series([None]*len(df))
        
    col_costo = []
    for c in cols:
        if 'COSTO FINAL' in c and 'KILO' not in c and 'PESOS' in c:
            col_costo.append(c)
            break
    if not col_costo:
        for c in cols:
            if 'COSTO' in c and 'PESOS' in c and 'KILO' not in c:
                col_costo.append(c)
                break
    if col_costo:
        mapped_df['costo_final_unitario'] = df[col_costo[0]]
    else:
        mapped_df['costo_final_unitario'] = pd.Series([None]*len(df))
        
    col_costo_kilo = []
    for c in cols:
        if 'KILO' in c and 'COSTO' in c and 'PESOS' in c:
            col_costo_kilo.append(c)
            break
    if not col_costo_kilo:
        for c in cols:
            if 'KILO' in c and 'COSTO' in c and not any(x in c for x in ['USD', 'FOB', 'FLETE']):
                col_costo_kilo.append(c)
                break
    if not col_costo_kilo:
        for c in cols:
            if 'KILO' in c and ('COSTO' in c or 'PRECIO' in c or 'VALOR' in c):
                col_costo_kilo.append(c)
                break
    if col_costo_kilo:
        mapped_df['costo_por_kilo'] = df[col_costo_kilo[0]]
    else:
        mapped_df['costo_por_kilo'] = pd.Series([None]*len(df))

    mapped_df['proveedor'] = df['_proveedor'] if '_proveedor' in df.columns else 'DESCONOCIDO'
    mapped_df['archivo_origen'] = df['_archivo'] if '_archivo' in df.columns else ''
    
    mapped_df['fecha'] = mapped_df['fecha'].ffill()
    mapped_df['nro_fc'] = mapped_df['nro_fc'].ffill()
    
    mapped_df['fecha'] = pd.to_datetime(mapped_df['fecha'], errors='coerce')
    mapped_df['kg_unitario'] = pd.to_numeric(mapped_df['kg_unitario'], errors='coerce').fillna(0.0)
    mapped_df['cantidad'] = pd.to_numeric(mapped_df['cantidad'], errors='coerce').fillna(0.0)
    mapped_df['costo_final_unitario'] = pd.to_numeric(mapped_df['costo_final_unitario'], errors='coerce').fillna(0.0)
    mapped_df['costo_por_kilo'] = pd.to_numeric(mapped_df['costo_por_kilo'], errors='coerce').fillna(0.0)
    
    mapped_df['guia'] = mapped_df['guia'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    mapped_df.loc[mapped_df['guia'].isin(['nan', 'None', '']), 'guia'] = '-'
    mapped_df['nro_fc'] = mapped_df['nro_fc'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    mapped_df.loc[mapped_df['nro_fc'].isin(['nan', 'None', '']), 'nro_fc'] = '-'
    
    def parse_article(desc):
        if not desc or pd.isna(desc):
            return "S/D", "S/D"
        desc = " ".join(str(desc).split()).strip().upper()
        parts = desc.split(maxsplit=1)
        if len(parts) == 2:
            return parts[0], parts[1]
        return parts[0], "SIN ESPECIFICAR"
        
    articles_split = mapped_df['articulo_original'].apply(parse_article)
    mapped_df['articulo_producto'] = [x[0] for x in articles_split]
    mapped_df['articulo_variedad'] = [x[1] for x in articles_split]
    
    mapped_df['kilos_totales'] = mapped_df['kg_unitario'] * mapped_df['cantidad']
    mapped_df['costo_final_total'] = mapped_df['costo_final_unitario'] * mapped_df['cantidad']
    
    mask_cero_kilo = (mapped_df['costo_por_kilo'] == 0.0) & (mapped_df['kg_unitario'] > 0.0)
    mapped_df.loc[mask_cero_kilo, 'costo_por_kilo'] = mapped_df['costo_final_unitario'] / mapped_df['kg_unitario']
    
    mapped_df = mapped_df.dropna(subset=['articulo_original'])
    return mapped_df

def read_excel_in_memory(uploaded_file):
    filename = uploaded_file.name
    sheet_dfs = []
    
    if filename.endswith('.xls'):
        import xlrd
        try:
            file_bytes = uploaded_file.read()
            book = xlrd.open_workbook(file_contents=file_bytes)
            for sheet_name in book.sheet_names():
                sheet = book.sheet_by_name(sheet_name)
                if sheet.nrows <= 1:
                    continue
                headers = [clean_header_name(sheet.cell_value(0, col)) for col in range(sheet.ncols)]
                sheet_rows = []
                for r in range(1, sheet.nrows):
                    row_dict = {'_proveedor': sheet_name.strip(), '_archivo': filename}
                    for c in range(min(len(headers), sheet.ncols)):
                        col_name = headers[c]
                        if not col_name:
                            continue
                        val = sheet.cell_value(r, c)
                        ctype = sheet.cell_type(r, c)
                        if ctype == xlrd.XL_CELL_DATE:
                            try:
                                date_tuple = xlrd.xldate_as_tuple(val, book.datemode)
                                val = datetime(*date_tuple)
                            except:
                                pass
                        elif ctype == xlrd.XL_CELL_EMPTY or val == '':
                            val = None
                        row_dict[col_name] = val
                    if any(v is not None for k, v in row_dict.items() if k not in ['_proveedor', '_archivo']):
                        sheet_rows.append(row_dict)
                if sheet_rows:
                    sheet_df = pd.DataFrame(sheet_rows)
                    cleaned_df = map_and_clean_df(sheet_df)
                    if not cleaned_df.empty:
                        sheet_dfs.append(cleaned_df)
        except Exception as e:
            st.error(f"Error en archivo XLS {filename}: {e}")
    else:
        import openpyxl
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                if sheet.max_row <= 1:
                    continue
                headers = [clean_header_name(cell.value) for cell in sheet[1]]
                sheet_rows = []
                for r in range(2, sheet.max_row + 1):
                    row_dict = {'_proveedor': sheet_name.strip(), '_archivo': filename}
                    for c in range(1, len(headers) + 1):
                        col_name = headers[c-1]
                        if not col_name:
                            continue
                        cell = sheet.cell(row=r, column=c)
                        val = cell.value
                        if val == '':
                            val = None
                        row_dict[col_name] = val
                    if any(v is not None for k, v in row_dict.items() if k not in ['_proveedor', '_archivo']):
                        sheet_rows.append(row_dict)
                if sheet_rows:
                    sheet_df = pd.DataFrame(sheet_rows)
                    cleaned_df = map_and_clean_df(sheet_df)
                    if not cleaned_df.empty:
                        sheet_dfs.append(cleaned_df)
        except Exception as e:
            st.error(f"Error en archivo XLSX {filename}: {e}")
    return sheet_dfs

# Función para convertir dataframe filtrado a Excel estilizado
def convert_df_to_excel(df_to_export):
    output = io.BytesIO()
    
    # Seleccionar y renombrar las columnas para la exportación final (con el orden exacto deseado)
    cols_to_export = [
        'fecha', 'proveedor', 'nro_fc', 'guia', 
        'articulo_producto', 'articulo_variedad',
        'kg_unitario', 'cantidad', 'costo_final_unitario',
        'costo_por_kilo', 'kilos_totales', 'costo_final_total'
    ]
    
    # Asegurar que existan todas las columnas
    available_cols = [c for c in cols_to_export if c in df_to_export.columns]
    df_sub = df_to_export[available_cols].copy()
    
    # Renombrar columnas a formato amigable (exactamente como en pantalla)
    col_names_map = {
        'fecha': 'Fecha',
        'proveedor': 'Proveedor',
        'nro_fc': 'Factura',
        'guia': 'Guía/Lote',
        'articulo_producto': 'Producto',
        'articulo_variedad': 'Variedad',
        'kg_unitario': 'Kg/Bulto',
        'cantidad': 'Bultos',
        'costo_final_unitario': 'Costo Final Unit. ($)',
        'costo_por_kilo': 'Costo por Kilo ($)',
        'kilos_totales': 'Total Kilos',
        'costo_final_total': 'Costo Final Total ($)'
    }
    df_sub = df_sub.rename(columns=col_names_map)
    
    # Formatear la columna fecha a string YYYY-MM-DD
    if 'Fecha' in df_sub.columns:
        df_sub['Fecha'] = df_sub['Fecha'].dt.strftime('%Y-%m-%d')
        
    # Crear archivo Excel en memoria usando xlsxwriter para aplicar estilos
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df_sub.to_excel(writer, sheet_name='Arribos Filtrados', index=False)
    
    workbook  = writer.book
    worksheet = writer.sheets['Arribos Filtrados']
    
    # Configurar formato de cabecera (Azul premium con texto blanco)
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'vcenter',
        'align': 'center',
        'fg_color': '#1E3A8A', # Azul marino
        'font_color': 'white',
        'border': 1
    })
    
    # Formatos de número
    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    weight_format = workbook.add_format({'num_format': '#,##0.00 "kg"'})
    qty_format = workbook.add_format({'num_format': '#,##0'})
    
    # Aplicar formato de cabecera a toda la fila 0
    for col_num, col_name in enumerate(df_sub.columns):
        worksheet.write(0, col_num, col_name, header_format)
        
        # Ajustar ancho de columna automáticamente de forma segura contra nulos
        non_null_vals = df_sub[col_name].dropna().astype(str)
        if not non_null_vals.empty:
            max_val_len = non_null_vals.map(len).max()
        else:
            max_val_len = 0
        max_len = max(max_val_len, len(col_name)) + 4
        worksheet.set_column(col_num, col_num, max_len)
        
        # Aplicar formato numérico a las columnas específicas
        if 'Costo' in col_name or '$' in col_name:
            worksheet.set_column(col_num, col_num, max_len, currency_format)
        elif 'Kilos' in col_name or 'Kg' in col_name:
            worksheet.set_column(col_num, col_num, max_len, weight_format)
        elif 'Cantidad' in col_name:
            worksheet.set_column(col_num, col_num, max_len, qty_format)
            
    writer.close()
    processed_data = output.getvalue()
    return processed_data

# Función para controlar el acceso (Login Screen)
def check_password():
    """Retorna True si el usuario ingresó la contraseña correcta."""
    correct_password = st.secrets.get("PASSWORD", None)
    
    # Si no se configuró contraseña en Secrets (ej: localmente), omitir la autenticación
    if not correct_password:
        return True

    if "password_correct" not in st.session_state:
        st.session_state.password_correct = False
        
    if st.session_state.password_correct:
        return True

    # Renderizar pantalla de login
    st.write("")
    st.write("")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style='background-color: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05); border: 1px solid #e2e8f0; text-align: center;'>
            <h2 style='color: #1e293b; font-family: sans-serif; margin-bottom: 5px; font-size: 24px;'>🔐 Acceso Protegido</h2>
            <p style='color: #64748b; font-size: 14px; margin-bottom: 10px;'>Ingresa la contraseña para acceder al Dashboard de Arribos</p>
        </div>
        """, unsafe_allow_html=True)
        st.write("")
        password = st.text_input("Contraseña de Acceso", type="password", key="login_password_input")
        if st.button("Ingresar al Tablero", use_container_width=True):
            if password == correct_password:
                st.session_state.password_correct = True
                st.rerun()
            else:
                st.error("Contraseña incorrecta. Por favor intenta nuevamente.")
                
    return False

# Validar credenciales
if not check_password():
    st.stop()

# Título Principal
st.title("📊 Dashboard de Próximos Arribos")
st.markdown("Consolidado e interactivo de importación de mercadería")
st.markdown("---")

# Cargar Datos (lectura en tiempo real de Google Sheets con fallback local)
gsheet_url = st.secrets.get("GSHEET_CSV_URL", None)
csv_mtime = os.path.getmtime(DATA_PATH) if os.path.exists(DATA_PATH) else 0
df = load_data(gsheet_url, csv_mtime)

# ------------------ SIDEBAR / PANEL DE CONTROL ------------------
st.sidebar.header("🔍 Panel de Control")

# Subidor de Archivos Excel integrado en la Web
with st.sidebar.expander("📤 Actualizar Base de Datos (Subir Excel)", expanded=False):
    uploaded_files = st.file_uploader(
        "Selecciona archivos .xls o .xlsx",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        key="excel_uploader"
    )
    
    if uploaded_files:
        if st.button("🚀 Procesar y Cargar a la Nube", use_container_width=True):
            new_dfs = []
            with st.spinner("Procesando archivos subidos..."):
                for file in uploaded_files:
                    dfs = read_excel_in_memory(file)
                    new_dfs.extend(dfs)
            
            if not new_dfs:
                st.warning("No se encontraron registros válidos para cargar.")
            else:
                new_data_df = pd.concat(new_dfs, ignore_index=True)
                
                # Control de Duplicados
                current_df = df.copy() if df is not None else pd.DataFrame()
                
                if not current_df.empty:
                    # Crear una clave única de coincidencia en base a los datos existentes
                    current_keys = set()
                    for _, row in current_df.iterrows():
                        f_str = pd.to_datetime(row['fecha']).strftime('%Y-%m-%d') if not pd.isna(row['fecha']) else ''
                        fc = str(row['nro_fc']).strip()
                        guia = str(row['guia']).strip()
                        prod = str(row['articulo_producto']).strip().upper()
                        var = str(row['articulo_variedad']).strip().upper()
                        cant = float(row['cantidad'])
                        current_keys.add((f_str, fc, guia, prod, var, cant))
                        
                    # Filtrar únicamente los registros nuevos
                    rows_to_insert = []
                    for _, row in new_data_df.iterrows():
                        f_str = pd.to_datetime(row['fecha']).strftime('%Y-%m-%d') if not pd.isna(row['fecha']) else ''
                        fc = str(row['nro_fc']).strip()
                        guia = str(row['guia']).strip()
                        prod = str(row['articulo_producto']).strip().upper()
                        var = str(row['articulo_variedad']).strip().upper()
                        cant = float(row['cantidad'])
                        
                        key = (f_str, fc, guia, prod, var, cant)
                        if key not in current_keys:
                            rows_to_insert.append(row)
                            
                    if not rows_to_insert:
                        st.info("Todos los registros de los archivos subidos ya existen en la base de datos (0 filas nuevas cargadas).")
                    else:
                        rows_to_insert_df = pd.DataFrame(rows_to_insert)
                        with st.spinner("Guardando registros en la base de datos..."):
                            success = save_new_rows(rows_to_insert_df)
                            if success:
                                st.success(f"¡Carga exitosa! Se agregaron {len(rows_to_insert_df)} registros nuevos.")
                                st.cache_data.clear()
                                st.rerun()
                else:
                    # Si la base de datos está vacía, cargar todo
                    with st.spinner("Iniciando base de datos con nuevos registros..."):
                        success = save_new_rows(new_data_df)
                        if success:
                            st.success(f"¡Base de datos iniciada con éxito! Se cargaron {len(new_data_df)} registros.")
                            st.cache_data.clear()
                            st.rerun()

if df is None or df.empty:
    st.info("💡 La base de datos está vacía. Por favor, sube tu primera planilla Excel en la barra lateral para ver el tablero.")
    if st.button("🔄 Ejecutar Consolidación Local (Desarrollador)"):
        with st.spinner("Consolidando archivos locales..."):
            try:
                import sys
                sys.path.append(os.path.dirname(os.path.abspath(__file__)))
                from update_data import consolidate_data
                consolidate_data()
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.exception(e)
else:
    # ------------------ SIDEBAR / FILTROS ------------------
    st.sidebar.header("🔍 Filtros del Tablero")
    
    # Filtro de Fechas
    min_date = df['fecha'].min().date()
    max_date = df['fecha'].max().date()
    
    # Por defecto, el rango de arranque empieza el 1 de enero del año en curso
    current_year = datetime.now().year
    start_default = date(current_year, 1, 1)
    if start_default > max_date:
        start_default = min_date
    elif start_default < min_date:
        start_default = min_date
        
    st.sidebar.subheader("Rango de Fechas")
    date_range = st.sidebar.date_input(
        "Selecciona el período",
        value=(start_default, max_date),
        min_value=min_date,
        max_value=max_date
    )
    
    # Filtro de Productos (Artículo)
    productos_disponibles = sorted(df['articulo_producto'].unique())
    selected_products = st.sidebar.multiselect(
        "Productos",
        options=productos_disponibles,
        default=[] # Por defecto vacio -> significa todos
    )
    
    # Filtro de Variedad (dinámico según los productos seleccionados)
    if selected_products:
        variedades_disponibles = sorted(df[df['articulo_producto'].isin(selected_products)]['articulo_variedad'].unique())
    else:
        variedades_disponibles = sorted(df['articulo_variedad'].unique())
        
    selected_varieties = st.sidebar.multiselect(
        "Variedades",
        options=variedades_disponibles,
        default=[] # Por defecto vacio -> significa todas
    )
    
    # Filtro de Guía / Lote
    guias_disponibles = sorted(df['guia'].unique())
    selected_guias = st.sidebar.multiselect(
        "Nro. Guía / Lote",
        options=guias_disponibles,
        default=[] # Por defecto vacio -> significa todas
    )
    
    # Filtro de Proveedor
    proveedores_disponibles = sorted(df['proveedor'].unique())
    selected_proveedores = st.sidebar.multiselect(
        "Proveedores",
        options=proveedores_disponibles,
        default=[] # Por defecto vacio -> significa todos
    )

    # ------------------ APLICAR FILTROS ------------------
    filtered_df = df.copy()
    
    # 1. Filtro de Fechas
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[
            (filtered_df['fecha'].dt.date >= start_date) & 
            (filtered_df['fecha'].dt.date <= end_date)
        ]
    elif isinstance(date_range, date):
        filtered_df = filtered_df[filtered_df['fecha'].dt.date == date_range]
        
    # 2. Filtro de Productos
    if selected_products:
        filtered_df = filtered_df[filtered_df['articulo_producto'].isin(selected_products)]
        
    # 2.5 Filtro de Variedades
    if selected_varieties:
        filtered_df = filtered_df[filtered_df['articulo_variedad'].isin(selected_varieties)]
        
    # 3. Filtro de Guías (Si está vacío, no se filtra y se muestran todas)
    if selected_guias:
        filtered_df = filtered_df[filtered_df['guia'].isin(selected_guias)]
        
    # 4. Filtro de Proveedores
    if selected_proveedores:
        filtered_df = filtered_df[filtered_df['proveedor'].isin(selected_proveedores)]

    # ------------------ SECCIÓN DE MÉTRICAS (KPIs) ------------------
    col1, col2, col3 = st.columns(3)
    
    total_kilos = filtered_df['kilos_totales'].sum()
    total_costo = filtered_df['costo_final_total'].sum()
    costo_medio_kg = total_costo / total_kilos if total_kilos > 0 else 0.0
    
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">⚖️ Kilaje Total</div>
            <div class="metric-value">{total_kilos:,.2f} kg</div>
        </div>
        """, unsafe_allow_html=True)
        
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">💰 Costo Final Total</div>
            <div class="metric-value">$ {total_costo:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)
        
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-title">📈 Costo Medio por Kilo</div>
            <div class="metric-value">$ {costo_medio_kg:,.2f}</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)

    # ------------------ ACCIONES Y TABLA DE DETALLES ------------------
    col_title, col_btn = st.columns([3, 1])
    
    with col_title:
        st.subheader("📋 Detalle de Arribos Seleccionados")
        
    with col_btn:
        if not filtered_df.empty:
            # Generar Excel en memoria
            excel_data = convert_df_to_excel(filtered_df)
            
            st.download_button(
                label="📥 Exportar a Excel",
                data=excel_data,
                file_name=f"arribos_filtrados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.button("📥 Exportar a Excel (Sin datos)", disabled=True, use_container_width=True)

    # Mostrar la tabla interactiva
    if filtered_df.empty:
        st.info("No se encontraron registros con los filtros seleccionados.")
    else:
        # Seleccionar y ordenar las columnas exactamente como se muestra en la captura del usuario
        cols_to_display = [
            'fecha', 'proveedor', 'nro_fc', 'guia', 
            'articulo_producto', 'articulo_variedad',
            'kg_unitario', 'cantidad', 'costo_final_unitario',
            'costo_por_kilo', 'kilos_totales', 'costo_final_total'
        ]
        display_df = filtered_df[cols_to_display].copy()
        
        # Resetear el índice a 0-based para sincronizar correctamente con los índices de selección de Streamlit
        display_df = display_df.reset_index(drop=True)
        
        # Renombrar columnas para coincidir exactamente con los encabezados deseados
        display_df = display_df.rename(columns={
            'fecha': 'Fecha',
            'proveedor': 'Proveedor',
            'nro_fc': 'Factura',
            'guia': 'Guía/Lote',
            'articulo_producto': 'Producto',
            'articulo_variedad': 'Variedad',
            'kg_unitario': 'Kg/Bulto',
            'cantidad': 'Bultos',
            'costo_final_unitario': 'Costo Final Unit.',
            'costo_por_kilo': 'Costo por Kilo',
            'kilos_totales': 'Total Kilos',
            'costo_final_total': 'Costo Final Total'
        })
        
        # Obtener filas seleccionadas desde el session_state para colorear la fila completa
        selected_rows = []
        if "arribos_grid" in st.session_state:
            grid_state = st.session_state["arribos_grid"]
            if "selection" in grid_state and "rows" in grid_state["selection"]:
                selected_rows = grid_state["selection"]["rows"]
                
        # Función para aplicar estilos dinámicos a la fila completa y columnas específicas
        def style_rows(row):
            styles = [''] * len(row)
            if row.name in selected_rows:
                # Pintar toda la fila seleccionada con un color azul brillante y texto claro
                return ['background-color: #2563eb; color: white; font-weight: bold'] * len(row)
            
            # Si no está seleccionada, aplicar el color azul marino en las columnas de Costo Final Unit. (idx 8) y Costo por Kilo (idx 9)
            styles[8] = 'background-color: #1E3A8A; color: white; font-weight: bold'
            styles[9] = 'background-color: #1E3A8A; color: white; font-weight: bold'
            return styles
            
        styled_df = display_df.style.format({
            'Fecha': lambda x: x.strftime('%Y-%m-%d') if not pd.isna(x) else '',
            'Kg/Bulto': '{:.2f} kg',
            'Bultos': '{:.0f}',
            'Costo Final Unit.': '${:.2f}',
            'Costo por Kilo': '${:.2f}',
            'Total Kilos': '{:.2f} kg',
            'Costo Final Total': lambda x: f"{x:,.2f}".replace(",", "x").replace(".", ",").replace("x", ".") + " $"
        }).apply(style_rows, axis=1)
        
        st.dataframe(
            styled_df,
            hide_index=True,
            width="stretch",
            on_select="rerun",
            selection_mode="single-row",
            key="arribos_grid"
        )
        
    # Información del pie de página
    st.markdown("---")
    source_info = "Google Sheets (Base de datos en la Nube)" if gsheet_url else f"Archivo local ({datetime.fromtimestamp(os.path.getmtime(DATA_PATH)).strftime('%Y-%m-%d %H:%M:%S')})" if os.path.exists(DATA_PATH) else "Desconocido"
    st.caption(f"Origen de datos: {source_info}")
