import pandas as pd
import os
import shutil
import re

# --- 1. CONFIGURACIÓN ---
# POR FAVOR, MODIFICA ESTAS TRES LÍNEAS:

# 1. El archivo consolidado que creamos en el script anterior
ruta_excel_consolidado = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TD GUAFITA\2.INFOENTRADA\DATOS_FINALES\resultado_consolidado.xlsx"

# 2. La carpeta donde están TODAS las fotos originales (el "pool" de fotos)
#    (Ej: \\192.168.1.2\cenyt-proyectos\CEN-223_TOP DRIVE\FOTOS_CRUDAS)
ruta_fotos_origen = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TD GUAFITA\2.INFOENTRADA\DATOS_FINALES\FOTOS_TRATADAS"

# 3. La carpeta base donde quieres crear la nueva estructura (ej: ...\FOTOS_ORDENADAS)
carpeta_destino_base = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TD GUAFITA\2.INFOENTRADA\FOTOS_ORDENADAS"

# --- 2. FUNCIÓN AUXILIAR (Para limpiar nombres de carpetas) ---
def limpiar_nombre_carpeta(nombre):
    """
    Elimina caracteres ilegales que no pueden usarse en nombres
    de carpetas en Windows.
    """
    if pd.isna(nombre) or not nombre:
        return None
    # Convertimos a string por si acaso (ej. un número de tag)
    nombre_str = str(nombre).strip()
    # Eliminamos caracteres ilegales: \ / : * ? " < > |
    return re.sub(r'[\\/*?:"<>|]', '_', nombre_str)

# --- 3. LECTURA Y PREPARACIÓN ---
print(f"Iniciando el proceso de organización de fotos...")
print(f"Leyendo Excel: {ruta_excel_consolidado}")

try:
    df = pd.read_excel(ruta_excel_consolidado, dtype=str)
except FileNotFoundError:
    print(f"¡Error! No se encontró el archivo Excel en: {ruta_excel_consolidado}")
    exit()
except Exception as e:
    print(f"¡Error inesperado al leer el Excel! {e}")
    exit()

# Verificar que las carpetas de origen y destino existan
if not os.path.isdir(ruta_fotos_origen):
    print(f"¡Error! La carpeta de origen de fotos NO existe: {ruta_fotos_origen}")
    exit()

if not os.path.isdir(carpeta_destino_base):
    print(f"Advertencia: La carpeta de destino base no existe. Se creará en: {carpeta_destino_base}")
    os.makedirs(carpeta_destino_base)


# Verificar columnas necesarias
columnas_necesarias = ['Circuito', 'Estructura_Tag']
if not all(col in df.columns for col in columnas_necesarias):
    print(f"¡Error! Al Excel le faltan una o más columnas clave: 'Circuito' o 'Estructura_Tag'.")
    exit()

# Identificar automáticamente todas las columnas de fotos
foto_columnas = sorted([col for col in df.columns if col.startswith("Foto")])

if not foto_columnas:
    print("¡Error! No se encontraron columnas que empiecen con 'Foto' (ej: Foto1, Foto2) en el Excel.")
    exit()

print(f"Columnas de fotos identificadas: {', '.join(foto_columnas)}")


# --- 4. PROCESO PRINCIPAL: COPIAR Y ORDENAR ---

print("\n--- Iniciando copia y organización ---")
print(f"Fotos de origen: {ruta_fotos_origen}")
print(f"Fotos de destino: {carpeta_destino_base}\n")

# Contadores para el resumen
fotos_copiadas = 0
fotos_no_encontradas = 0
filas_omitidas = 0
set_fotos_no_encontradas = set()

# Iteramos sobre cada fila del DataFrame
for index, fila in df.iterrows():
    
    # 1. Obtener y limpiar nombres de carpetas
    nombre_circuito = limpiar_nombre_carpeta(fila['Circuito'])
    nombre_estructura = limpiar_nombre_carpeta(fila['Estructura_Tag'])
    
    # 2. Validar que tengamos los datos para crear la carpeta
    if not nombre_circuito or not nombre_estructura:
        print(f"  Fila {index+2} omitida: 'Circuito' o 'Estructura_Tag' están vacíos.")
        filas_omitidas += 1
        continue
        
    # 3. Crear la ruta de destino final
    ruta_destino_final = os.path.join(carpeta_destino_base, nombre_circuito, nombre_estructura)
    
    # 4. Crear las carpetas (si no existen)
    #    os.makedirs crea todas las carpetas intermedias (ej: Circuito y Estructura_Tag)
    try:
        os.makedirs(ruta_destino_final, exist_ok=True)
    except OSError as e:
        print(f"  Error creando carpeta {ruta_destino_final}. Omitiendo fila. Error: {e}")
        filas_omitidas += 1
        continue

    # 5. Iterar sobre las columnas de fotos (Foto1, Foto2...) para esta fila
    for col in foto_columnas:
        nombre_foto = fila[col]
        
        # Si la celda de la foto está vacía, la saltamos
        if pd.isna(nombre_foto) or not nombre_foto:
            continue
            
        nombre_foto = str(nombre_foto).strip() # Limpiar espacios
        
        # 6. Definir rutas de origen y destino del archivo
        ruta_origen_foto = os.path.join(ruta_fotos_origen, nombre_foto)
        ruta_destino_foto = os.path.join(ruta_destino_final, nombre_foto)
        
        # 7. Verificar si la foto existe en el ORIGEN
        if os.path.exists(ruta_origen_foto):
            # 8. Verificar si la foto YA existe en el DESTINO (para no copiarla de nuevo)
            if not os.path.exists(ruta_destino_foto):
                try:
                    # Usamos copy2 para intentar preservar metadatos (como fecha de creación)
                    shutil.copy2(ruta_origen_foto, ruta_destino_foto)
                    fotos_copiadas += 1
                except Exception as e:
                    print(f"  ¡Error copiando {nombre_foto}! Error: {e}")
            # else:
                # La foto ya existe en el destino, no hacemos nada.
        else:
            # La foto no se encontró en la carpeta de origen
            fotos_no_encontradas += 1
            set_fotos_no_encontradas.add(nombre_foto)

# --- 5. REPORTE FINAL ---
print("\n--- ¡Proceso completado! ---")
print(f"Fotos copiadas exitosamente: {fotos_copiadas}")
print(f"Filas omitidas (sin Circuito/Tag): {filas_omitidas}")
print(f"Fotos no encontradas en origen: {fotos_no_encontradas}")

if fotos_no_encontradas > 0:
    print("\nLista de algunas fotos no encontradas:")
    # Mostramos las primeras 20 fotos no encontradas
    for i, foto in enumerate(list(set_fotos_no_encontradas)):
        if i >= 20:
            print(f"  ...y {len(set_fotos_no_encontradas) - 20} más.")
            break
        print(f"  - {foto}")