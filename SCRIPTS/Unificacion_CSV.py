import pandas as pd
import glob
import os
import re # Importamos la librería de Expresiones Regulares (Regex)

# --- 1. CONFIGURACIÓN ---
# POR FAVOR, MODIFICA ESTAS DOS LÍNEAS:

# 1. La carpeta donde están TODOS tus archivos (CSV y Excel)
carpeta_entrada = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TOP DRIVE\2.INFOENTRADA\DATOS_CRUDOS\10-23-2025"

# 2. La ruta COMPLETA donde quieres guardar el Excel final (incluye el nombre.xlsx)
ruta_excel_salida = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TOP DRIVE\2.INFOENTRADA\DATOS_FINALES\resultado_consolidado.xlsx"

# --------------------------

# --- 2. FUNCIÓN AUXILIAR (Conversión DMS a DD) ---
# Esta función convierte el formato 7° 2'30.97"N a 7.0419...

def dms_to_dd(dms_str):
    """
    Convierte una cadena de Grados, Minutos, Segundos (DMS) a Grados Decimales (DD).
    """
    # Si el valor es nulo o no es texto, devuelve nulo
    if pd.isna(dms_str) or not isinstance(dms_str, str):
        return pd.NA

    # Expresión regular para extraer Grados, Minutos, Segundos y Hemisferio
    pattern = re.compile(r"(\d+)\s*°\s*(\d+)\s*'\s*([\d\.]+)\"\s*([NSOEW])")
    match = pattern.search(dms_str.strip())

    # Si el formato no coincide, devuelve nulo
    if not match:
        return pd.NA

    try:
        degrees = float(match.group(1))
        minutes = float(match.group(2))
        seconds = float(match.group(3))
        hemisphere = match.group(4).upper() # N, S, O, E, W

        # Fórmula de conversión
        dd = degrees + (minutes / 60) + (seconds / 3600)

        # Aplicar signo negativo para Sur u Oeste (Oeste)
        if hemisphere in ['S', 'O', 'W']:
            dd = -dd
            
        return dd
    except Exception:
        # Si algo falla en la conversión
        return pd.NA

# --------------------------

# --- 3. BÚSQUEDA Y LECTURA DE ARCHIVOS ---
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


    # --- 4. CONVERSIÓN DE COORDENADAS (DMS a DD) ---
    # --- LÓGICA CORREGIDA ---
    print("Iniciando conversión de coordenadas GPS (DMS a DD)...")
    
    # Verificar si existen las columnas de GPS
    if 'Latitud_GPS' in df_final.columns and 'Longitud_GPS' in df_final.columns:
        
        # 1. Calcular los valores de GPS (se usarán solo para rellenar)
        print("  Calculando valores desde Latitud_GPS y Longitud_GPS...")
        lat_calculada = df_final['Latitud_GPS'].apply(dms_to_dd)
        lon_calculada = df_final['Longitud_GPS'].apply(dms_to_dd)

        # 2. Asegurar que las columnas Manual existan (si no, las crea vacías)
        if 'Latitud_Manual' not in df_final.columns:
            df_final['Latitud_Manual'] = pd.NA
        if 'Longitud_Manual' not in df_final.columns:
            df_final['Longitud_Manual'] = pd.NA

        # 3. Reemplazar celdas con texto vacío ('') con NaN.
        #    Esto asegura que fillna() detecte tanto las celdas NaN como las vacías.
        #    Usamos .loc para evitar advertencias de "SettingWithCopyWarning"
        df_final.loc[df_final['Latitud_Manual'] == '', 'Latitud_Manual'] = pd.NA
        df_final.loc[df_final['Longitud_Manual'] == '', 'Longitud_Manual'] = pd.NA
        
        # 4. Rellenar *solo* los vacíos (NaN) en las columnas Manual
        #    usando los valores calculados de GPS.
        #    Si ya hay un valor en Latitud_Manual, fillna() lo ignorará.
        df_final['Latitud_Manual'] = df_final['Latitud_Manual'].fillna(lat_calculada)
        df_final['Longitud_Manual'] = df_final['Longitud_Manual'].fillna(lon_calculada)
        
        print("Conversión completada. (Se respetaron los datos manuales existentes)")
    else:
        print("Advertencia: No se encontraron las columnas 'Latitud_GPS' o 'Longitud_GPS'. Saltando conversión de coordenadas.")
    
    
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
        
        df_fotos_separadas = df_final[columna_objetivo].str.split(',', expand=True)
        
        nuevos_nombres = {i: f"Foto{i+1}" for i in df_fotos_separadas.columns}
        df_fotos_separadas = df_fotos_separadas.rename(columns=nuevos_nombres)
        
        print(f"Se crearon {len(df_fotos_separadas.columns)} columnas de fotos.")
        
        # --- 7. UNIR Y GUARDAR ---
        
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