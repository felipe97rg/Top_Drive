import pandas as pd
import glob
import os
import re # Importamos la librería de Expresiones Regulares (Regex)

# --- 1. CONFIGURACIÓN ---
# POR FAVOR, MODIFICA ESTAS DOS LÍNEAS:

# 1. La carpeta donde están TODOS tus archivos (CSV y Excel)
carpeta_entrada = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TOP DRIVE\2.INFOENTRADA\DATOS_CRUDOS\10-25-2025"

# 2. La ruta COMPLETA donde quieres guardar el Excel final (incluye el nombre.xlsx)
ruta_excel_salida = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TOP DRIVE\2.INFOENTRADA\DATOS_FINALES\resultado_consolidado.xlsx"

# --------------------------

# --- 2. FUNCIÓN AUXILIAR (Conversión Decimal con Comas) ---
# Esta función convierte formatos como "7,033,365" o "71.01" a 7.033365 o 71.01

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

# --- 3. BÚSQUEDA Y LECTURA DE ARCHIVOS (CONCATENACIÓN) ---
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
            df_temp = pd.read_csv(archivo, sep=';', dtype=str, encoding='utf-8')
            lista_dataframes.append(df_temp)
        except Exception as e:
            print(f"  Advertencia: No se pudo leer el CSV '{os.path.basename(archivo)}'. Error: {e}")

    # --- Procesar Excels ---
    print("Leyendo archivos Excel...")
    for archivo in lista_archivos_excel:
        try:
            df_temp = pd.read_excel(archivo, dtype=str)
            lista_dataframes.append(df_temp)
        except Exception as e:
            print(f"  Advertencia: No se pudo leer el Excel '{os.path.basename(archivo)}'. Error: {e}")

    # --- Concatenación ---
    print("Concatenando todos los archivos...")
    df_final = pd.concat(lista_dataframes, ignore_index=True)
    print("¡Archivos concatenados!")


    # --- 4. CORRECCIÓN DE COORDENADAS Y RELLENO ---
    print("Iniciando limpieza y conversión de coordenadas...")

    # Paso 1: Asegurar que las columnas Manual existan
    if 'Latitud_Manual' not in df_final.columns:
        df_final['Latitud_Manual'] = pd.NA
    if 'Longitud_Manual' not in df_final.columns:
        df_final['Longitud_Manual'] = pd.NA

    # Verificar si existen las columnas de GPS
    if 'Latitud_GPS' in df_final.columns and 'Longitud_GPS' in df_final.columns:
        
        # Paso 2: Corregir las 4 columnas (convertirlas a numérico)
        print("  Convirtiendo columnas GPS a formato numérico...")
        df_final['Latitud_GPS'] = df_final['Latitud_GPS'].apply(convertir_a_decimal)
        df_final['Longitud_GPS'] = df_final['Longitud_GPS'].apply(convertir_a_decimal)
        
        print("  Convirtiendo columnas Manuales a formato numérico...")
        df_final['Latitud_Manual'] = df_final['Latitud_Manual'].apply(convertir_a_decimal)
        df_final['Longitud_Manual'] = df_final['Longitud_Manual'].apply(convertir_a_decimal)

        # Paso 3: Llenar los vacíos
        # (Rellenar Latitud_Manual (que ya es numérica) con Latitud_GPS (que ya es numérica))
        print("  Rellenando vacíos de columnas Manuales con datos de GPS...")
        df_final['Latitud_Manual'] = df_final['Latitud_Manual'].fillna(df_final['Latitud_GPS'])
        df_final['Longitud_Manual'] = df_final['Longitud_Manual'].fillna(df_final['Longitud_GPS'])
        
        print("Conversión y relleno completados.")
    else:
        print("Advertencia: No se encontraron las columnas 'Latitud_GPS' o 'Longitud_GPS'. Saltando relleno de coordenadas.")
    
    
    # --- 5. TRANSFORMACIÓN DE TEXTO (Ruta_Fotos) ---
    columna_objetivo = "Ruta_Fotos"
    texto_a_eliminar = "/storage/emulated/0/Download/SALI/"

    if columna_objetivo not in df_final.columns:
        print(f"Advertencia: La columna '{columna_objetivo}' no existe. Saltando limpieza de rutas de fotos.")
    else:
        print(f"Limpiando la columna '{columna_objetivo}'...")
        
        df_final[columna_objetivo] = df_final[columna_objetivo].fillna('').astype(str)
        df_final[columna_objetivo] = df_final[columna_objetivo].str.replace(texto_a_eliminar, "")
        
        # --- 6. SEPARACIÓN EN COLUMNAS (Fotos) ---
        print("Separando rutas en columnas 'Foto1', 'Foto2', ...")
        
        # Asegurarnos de que el separador sea una coma
        df_fotos_separadas = df_final[columna_objetivo].str.split(',', expand=True)
        
        nuevos_nombres = {i: f"Foto{i+1}" for i in df_fotos_separadas.columns}
        df_fotos_separadas = df_fotos_separadas.rename(columns=nuevos_nombres)
        
        print(f"Se crearon {len(df_fotos_separadas.columns)} columnas de fotos.")
        
        # --- 7. UNIR Y GUARDAR (CREAR ARCHIVO DE SALIDA) ---
        
        # Unimos el DataFrame original (df_final) con las nuevas columnas de fotos
        df_resultado = df_final.join(df_fotos_separadas)
        
        # Opcional: Eliminar la columna "Rutas_Fotos" original
        # if columna_objetivo in df_resultado.columns:
        #     df_resultado = df_resultado.drop(columns=[columna_objetivo])
        
        print(f"Guardando archivo Excel en: {ruta_excel_salida}")
        try:
            # Guardar en Excel sin el índice de pandas
            df_resultado.to_excel(ruta_excel_salida, index=False)
            print("\n¡Proceso completado exitosamente!")
            
        except Exception as e:
            print(f"¡Error al guardar el Excel! Verifica la ruta y permisos: {e}")