import pandas as pd
import glob
import os
import re # Importamos la librería de Expresiones Regulares (Regex)
import numpy as np # Importamos numpy para la lógica condicional

# --- 1. CONFIGURACIÓN ---
# POR FAVOR, MODIFICA ESTAS DOS LÍNEAS:

# 1. La carpeta donde están TODOS tus archivos (CSV y Excel)
carpeta_entrada = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TD GUAFITA\2.INFOENTRADA\DATOS_CRUDOS\Tablas"

# 2. La ruta COMPLETA donde quieres guardar el Excel final (incluye el nombre.xlsx)
ruta_excel_salida = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TD GUAFITA\2.INFOENTRADA\DATOS_FINALES\resultado_consolidado.xlsx"

# --- 2. FUNCIÓN AUXILIAR (Conversión Decimal con Comas) ---
def convertir_a_decimal(valor):
    """
    Convierte una cadena que usa comas como decimales/separadores
    a un número flotante (Grados Decimales).
    """
    # 1. Si es nulo o ya es un número, devolverlo
    if pd.isna(valor) or isinstance(valor, (int, float)):
        return valor
    
    # 2. Convertir a string y limpiar espacios
    try:
        s = str(valor).strip()
    except Exception:
        return pd.NA
    
    # Si la cadena está vacía después de limpiarla
    if not s:
        return pd.NA

    # 3. Si ya tiene un punto decimal (ej: "71.01"), 
    #    solo limpiamos las comas (si las hay) y convertimos
    if '.' in s:
        s = s.replace(',', '')
        return pd.to_numeric(s, errors='coerce')

    # 4. Si NO tiene punto, asumimos que la PRIMERA coma es el decimal
    #    y las demás comas son ruido (separadores de miles).
    
    # Reemplazar la primera coma por un punto
    s = s.replace(',', '.', 1)
    
    # Eliminar todas las comas restantes
    s = s.replace(',', '')

    # 5. Intentar convertir a número
    return pd.to_numeric(s, errors='coerce')

# --------------------------

# --- 3. FUNCIÓN AUXILIAR (Corrección de Codificación) ---
def corregir_codificacion(df):
    """
    Corrige los errores comunes de codificación (mojibake) en todo el DataFrame.
    """
    print("Corrigiendo errores de codificación (ej: 'Ã­' por 'í')...")
    
    # Mapeo de los errores de codificación más comunes
    replacements = {
        'Ã¡': 'á', 'Ã©': 'é', 'Ã­':'í', 'Ã³': 'ó', 'Ãº': 'ú', 'Ã±': 'ñ',
        'Ã': 'Á', 'Ã‰': 'É', 'Ã': 'Í', 'Ã“': 'Ó', 'Ãš': 'Ú', 'Ã‘': 'Ñ',
        'Â': ''  # A veces aparece un carácter 'Â' extra
    }
    
    # Seleccionamos solo las columnas de tipo 'object' (strings)
    columnas_str = df.select_dtypes(include=['object']).columns
    
    # Aplicamos el reemplazo usando regex para todas las ocurrencias
    for col in columnas_str:
        # Usamos .astype(str) para manejar datos mixtos (ej. None, NaN)
        df[col] = df[col].astype(str).replace(replacements, regex=True)
        
    print("Corrección de codificación completada.")
    return df

# --------------------------

# --- 4. BÚSQUEDA Y LECTURA DE ARCHIVOS ---
print(f"Buscando archivos en: {carpeta_entrada}")

patron_csv = os.path.join(carpeta_entrada, "*.csv")
patron_excel_xlsx = os.path.join(carpeta_entrada, "*.xlsx")
patron_excel_xls = os.path.join(carpeta_entrada, "*.xls")

lista_archivos_csv = glob.glob(patron_csv)
lista_archivos_excel = glob.glob(patron_excel_xlsx) + glob.glob(patron_excel_xls)

if not lista_archivos_csv and not lista_archivos_excel:
    print("¡Error! No se encontraron archivos CSV ni Excel en la carpeta especificada.")
else:
    print(f"Se encontraron {len(lista_archivos_csv)} archivos CSV.")
    print(f"Se encontraron {len(lista_archivos_excel)} archivos Excel.")
    
    lista_dataframes = []

    # --- Procesar CSVs ---
    print("Leyendo archivos CSV...")
    for archivo in lista_archivos_csv:
        try:
            # Leemos como UTF-8 por defecto
            df_temp = pd.read_csv(archivo, sep=';', dtype=str, encoding='utf-8')
            lista_dataframes.append(df_temp)
        except UnicodeDecodeError:
            try:
                # Si falla UTF-8, intentamos con 'latin-1' (común en Windows)
                df_temp = pd.read_csv(archivo, sep=';', dtype=str, encoding='latin-1')
                lista_dataframes.append(df_temp)
                print(f"   Advertencia: CSV '{os.path.basename(archivo)}' leído como 'latin-1'.")
            except Exception as e:
                print(f"   Advertencia: No se pudo leer el CSV '{os.path.basename(archivo)}'. Error: {e}")
        except Exception as e:
             print(f"   Advertencia: No se pudo leer el CSV '{os.path.basename(archivo)}'. Error: {e}")

    # --- Procesar Excels ---
    print("Leyendo archivos Excel...")
    for archivo in lista_archivos_excel:
        try:
            df_temp = pd.read_excel(archivo, dtype=str)
            lista_dataframes.append(df_temp)
        except Exception as e:
            print(f"   Advertencia: No se pudo leer el Excel '{os.path.basename(archivo)}'. Error: {e}")

    # --- 5. CONCATENACIÓN Y LIMPIEZA INICIAL ---
    print("Concatenando todos los archivos...")
    df_final = pd.concat(lista_dataframes, ignore_index=True)
    print("¡Archivos concatenados!")
    
    # Aplicamos la corrección de codificación aquí
    df_final = corregir_codificacion(df_final)
    
    # Reemplazamos 'nan' (como string, que viene de la corrección) por vacío
    df_final = df_final.replace('nan', pd.NA)

    # --- 5.1. (NUEVO) REEMPLAZO DE VALORES EN 'Circuito' ---
    print("\nIniciando reemplazos en la columna 'Circuito'...")
    if 'Circuito' in df_final.columns:
        # Primero, limpiamos espacios en blanco para asegurar coincidencias exactas
        df_final['Circuito'] = df_final['Circuito'].astype(str).str.strip()

        # Definimos los reemplazos que solicitaste
        reemplazos_circuito = {
            'Q2- Derivación': 'Q2',           # Interpretando "Q2- Derivación esto Q2"
            'Q6 derivación con Q2': 'Q2',     # Literal de la solicitud
            'Q6 derivación CTO 2 almacen': 'Q6' # Literal de la solicitud
        }
        
        # Aplicar los reemplazos (coincidencia exacta del valor de la celda)
        df_final['Circuito'] = df_final['Circuito'].replace(reemplazos_circuito)
        
        print("Reemplazos en 'Circuito' completados.")
    else:
        print("Advertencia: No se encontró la columna 'Circuito'. Saltando reemplazos.")
    # --------------------------------------------------


    # --- 6. CORRECCIÓN DE COORDENADAS Y RELLENO ---
    print("\nIniciando limpieza y conversión de coordenadas...")

    if 'Latitud_Manual' not in df_final.columns:
        df_final['Latitud_Manual'] = pd.NA
    if 'Longitud_Manual' not in df_final.columns:
        df_final['Longitud_Manual'] = pd.NA

    if 'Latitud_GPS' in df_final.columns and 'Longitud_GPS' in df_final.columns:
        print("   Convirtiendo columnas GPS a formato numérico...")
        df_final['Latitud_GPS'] = df_final['Latitud_GPS'].apply(convertir_a_decimal)
        df_final['Longitud_GPS'] = df_final['Longitud_GPS'].apply(convertir_a_decimal)
        
        print("   Convirtiendo columnas Manuales a formato numérico...")
        df_final['Latitud_Manual'] = df_final['Latitud_Manual'].apply(convertir_a_decimal)
        df_final['Longitud_Manual'] = df_final['Longitud_Manual'].apply(convertir_a_decimal)

        print("   Rellenando vacíos de columnas Manuales con datos de GPS...")
        df_final['Latitud_Manual'] = df_final['Latitud_Manual'].fillna(df_final['Latitud_GPS'])
        df_final['Longitud_Manual'] = df_final['Longitud_Manual'].fillna(df_final['Longitud_GPS'])
        
        print("Conversión y relleno completados.")
    else:
        print("Advertencia: No se encontraron las columnas 'Latitud_GPS' o 'Longitud_GPS'. Saltando relleno de coordenadas.")
    
    
    # --- 7. TRANSFORMACIÓN DE TEXTO (Ruta_Fotos) ---
    columna_fotos = "Ruta_Fotos"
    texto_a_eliminar = "/storage/emulated/0/Download/SALI/"

    if columna_fotos not in df_final.columns:
        print(f"\nAdvertencia: La columna '{columna_fotos}' no existe. Saltando limpieza y separación de fotos.")
        # Si no hay fotos, el df_resultado es solo una copia de df_final
        df_resultado = df_final.copy()
    else:
        print(f"\nLimpiando la columna '{columna_fotos}'...")
        
        # .astype(str) por si acaso la columna quedó como NA
        df_final[columna_fotos] = df_final[columna_fotos].astype(str).str.replace(texto_a_eliminar, "")
        
        # --- 8. SEPARACIÓN EN COLUMNAS (Fotos) ---
        print("Separando rutas en columnas 'Foto1', 'Foto2', ...")
        
        df_fotos_separadas = df_final[columna_fotos].str.split(',', expand=True)
        
        nuevos_nombres = {i: f"Foto{i+1}" for i in df_fotos_separadas.columns}
        df_fotos_separadas = df_fotos_separadas.rename(columns=nuevos_nombres)
        
        print(f"Se crearon {len(df_fotos_separadas.columns)} columnas de fotos.")
        
        # --- 9. UNIR DATAFRAMES (Principal + Fotos) ---
        # Si hay fotos, df_resultado es la unión
        df_resultado = df_final.join(df_fotos_separadas)
        
        
    # --- 10. CREAR Y GUARDAR REPORTE DE HALLAZGOS (LÓGICA MOVIDA) ---
    print("\n--- Iniciando creación de reporte 'Hallazgos' ---")
    
    cols_requeridas_hallazgos = [
        'Circuito', 'Estructura_Tag', 'Apoyo_Fractura', 
        'Templetes_Rotos', 'Templetes_Faltantes',
        'Templetes_Flojos',
        'Templete_Observaciones'
    ]
    
    # Usamos df_resultado (que siempre existe)
    cols_faltantes = [col for col in cols_requeridas_hallazgos if col not in df_resultado.columns]

    if cols_faltantes:
        print(f"   Advertencia: No se puede crear 'Hallazgos.xlsx'. Faltan columnas: {cols_faltantes}")
    else:
        print("   Generando columnas para Hallazgos...")
        
        # 2. Crear 'Custom'
        # NOTA: 'Circuito' ya está limpio gracias al paso 5.1
        circuito_str = df_resultado['Circuito'].astype(str).str.strip().fillna('')
        tag_str = df_resultado['Estructura_Tag'].astype(str).str.strip().fillna('')
        df_resultado['Custom'] = circuito_str + " " + tag_str

        # 3. Crear 'REEMPLAZO DE POSTES'
        col_comparacion = df_resultado['Apoyo_Fractura'].astype(str).str.strip().str.lower()
        condicion_fractura = (col_comparacion == 'sí')
        df_resultado['REEMPLAZO DE POSTES'] = condicion_fractura.astype(int)
        
        # 4. Crear 'INSTALACION RETENIDAS NUEVAS'
        print("   Calculando 'INSTALACION RETENIDAS NUEVAS'...")
        col_rotos = pd.to_numeric(df_resultado['Templetes_Rotos'], errors='coerce').fillna(0)
        col_faltantes = pd.to_numeric(df_resultado['Templetes_Faltantes'], errors='coerce').fillna(0)
        suma_inicial = col_rotos + col_faltantes
        condicion_vacio = (suma_inicial == 0)
        obs_col = df_resultado['Templete_Observaciones'].astype(str).fillna('').str.lower()
        regex_inclinacion = r'inclinad[oa]|inclinacion'
        condicion_obs = obs_col.str.contains(regex_inclinacion, regex=True, na=False)
        df_resultado['INSTALACION RETENIDAS NUEVAS'] = np.where(
            (condicion_vacio & condicion_obs), 
            1, 
            suma_inicial
        )
        
        # 5. Crear 'RETENSIONADO RETENIDAS'
        print("   Calculando 'RETENSIONADO RETENIDAS'...")
        col_flojos = pd.to_numeric(df_resultado['Templetes_Flojos'], errors='coerce').fillna(0)
        df_resultado['RETENSIONADO RETENIDAS'] = col_flojos
        
        
        # --- (NUEVO) 11. CREACIÓN DE COLUMNAS DE NORMALIZACIÓN (LÓGICA MOVIDA AQUÍ) ---
        print("\n   --- Iniciando creación de columnas de Normalización (13 kv a 35 kv) ---")
        
        # Definir nombres de columnas
        col_norm_ret = 'NORMALIZACION ESTRUCTURA RETENCION 13 kv a 35 kv'
        col_norm_sus = 'NORMALIZACION ESTRUCTURA SUSPENSION 13 kv a 35 kv'
        
        cols_requeridas_norm = ['Aisladores_Observaciones', 'Configuracion']
        cols_faltantes_norm = [col for col in cols_requeridas_norm if col not in df_resultado.columns]
        
        # Lista para rastrear las columnas que SÍ creamos
        columnas_norm_creadas = []

        if cols_faltantes_norm:
            print(f"     Advertencia: No se pueden crear columnas de Normalización. Faltan: {cols_faltantes_norm}")
        else:
            print("     Calculando columnas de Normalización...")
            
            # Condición 1: Contiene "2 discos"
            cond_discos = df_resultado['Aisladores_Observaciones'].astype(str).str.contains(
                "2 discos", case=False, na=False
            )
            
            # Condición 2: El tipo de configuración
            config_limpia = df_resultado['Configuracion'].astype(str).str.strip()
            cond_retencion = (config_limpia == "Retención (Corte)")
            cond_suspension = (config_limpia == "Suspensión")

            # --- Crear las nuevas columnas con 1 y 0 ---
            
            # Columna de Retención
            df_resultado[col_norm_ret] = np.where(
                cond_discos & cond_retencion,
                1,  # (NUEVO)
                0   # (NUEVO)
            )

            # Columna de Suspensión
            df_resultado[col_norm_sus] = np.where(
                cond_discos & cond_suspension,
                1,  # (NUEVO)
                0   # (NUEVO)
            )
            
            # Añadir a la lista de "creadas" para el reporte
            columnas_norm_creadas = [col_norm_ret, col_norm_sus]
            print("     Columnas de Normalización creadas exitosamente.")
        
        # --- Fin de la sección de Normalización ---
        
        
        # 6. Seleccionar columnas para Hallazgos (ACTUALIZADO)
        print("\n   Seleccionando columnas finales para Hallazgos...")
        
        columnas_finales_hallazgos = [
            'Circuito', 'Estructura_Tag', 'Custom', 
            'REEMPLAZO DE POSTES', 
            'INSTALACION RETENIDAS NUEVAS',
            'RETENSIONADO RETENIDAS'
        ]
        
        # (NUEVO) Añadir las columnas de normalización (si se crearon)
        columnas_finales_hallazgos.extend(columnas_norm_creadas)
        
        df_hallazgos = df_resultado[columnas_finales_hallazgos]
        
        # 7. Definir la ruta de salida
        directorio_salida = os.path.dirname(ruta_excel_salida)
        ruta_hallazgos_salida = os.path.join(directorio_salida, "Hallazgos.xlsx")
        
        # 8. Guardar el nuevo Excel
        try:
            print(f"   Guardando archivo de Hallazgos en: {ruta_hallazgos_salida}")
            df_hallazgos.to_excel(ruta_hallazgos_salida, index=False)
            print("   Archivo 'Hallazgos.xlsx' guardado exitosamente.")
        except Exception as e_hallazgos:
            print(f"   ¡Error al guardar el Excel 'Hallazgos'! Verifica la ruta y permisos: {e_hallazgos}")
    
    # --------------------------------------------------------

    # --- 12. GUARDAR ARCHIVO DE SALIDA PRINCIPAL ---
    # (Ya no hay sección 11, esta es la 12)
    # Nota: df_resultado ya tiene las columnas de normalización
    
    print(f"\nGuardando archivo principal en: {ruta_excel_salida}")
    try:
        df_resultado.to_excel(ruta_excel_salida, index=False)
        print("\n¡Proceso completado exitosamente!")
        
    except Exception as e:
        print(f"¡Error al guardar el Excel principal! Verifica la ruta y permisos: {e}")