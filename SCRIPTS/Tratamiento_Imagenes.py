import os
import shutil
from PIL import Image

# --- Rutas (Las que actualizaste) ---
carpeta = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TD GUAFITA\2.INFOENTRADA\DATOS_FINALES/FOTOS_TRATADAS"
carpeta_salida = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TD GUAFITA\2.INFOENTRADA\DATOS_FINALES/FOTOS_TRATADAS"
os.makedirs(carpeta_salida, exist_ok=True)
# Límite en KB
limite_kb = 500

# --- FUNCIÓN "procesar_imagen" (SIN CAMBIOS) ---
# Esta función está perfecta, la dejamos como está.
def procesar_imagen(ruta, destino, limite_kb, calidad_inicial=85, paso=5, factor_resize=0.9):
    """
    Procesa una imagen:
    1. Rota 90 grados a la derecha si es más ancha que alta.
    2. Convierte PNG a JPG.
    3. Comprime si supera el 'limite_kb'.
    4. Guarda la imagen en el destino.
    """
    try:
        img = Image.open(ruta)
        tamano_original_kb = os.path.getsize(ruta) / 1024
        
        # 1. Verificar dimensiones y rotar si es necesario
        ancho, alto = img.size
        fue_rotada = False
        if ancho > alto:
            img = img.transpose(Image.ROTATE_270)
            fue_rotada = True
            ancho, alto = img.size 

        # 2. Manejar PNG -> JPG
        if ruta.lower().endswith(".png"):
            img = img.convert("RGB")
            # El 'destino' ya viene con .jpg, pero nos aseguramos
            destino = os.path.splitext(destino)[0] + ".jpg"

        # 3. Decidir si comprimir o solo guardar
        
        # Caso A: No necesitaba cambios (y el destino es el mismo)
        if tamano_original_kb <= limite_kb and not fue_rotada and not ruta.lower().endswith(".png"):
            shutil.copy2(ruta, destino)
            print(f"   → Copiada sin cambios ({tamano_original_kb:.2f} KB)")
            return

        # Caso B: Fue rotada/convertida o estaba por debajo del límite
        if tamano_original_kb <= limite_kb:
            img.save(destino, quality=95, optimize=True)
            nuevo_tamano = os.path.getsize(destino) / 1024
            print(f"   → Guardada (Rotada/Convertida): {nuevo_tamano:.2f} KB")
            return

        # Caso C: Superaba el límite (necesita compresión)
        calidad = calidad_inicial
        while True:
            img.save(destino, optimize=True, quality=calidad)
            tamano_kb = os.path.getsize(destino) / 1024

            if tamano_kb <= limite_kb:
                print(f"   → Comprimida: {tamano_kb:.2f} KB")
                return 

            if calidad > 20:
                calidad -= paso
            else:
                ancho = int(ancho * factor_resize)
                alto = int(alto * factor_resize)
                img = img.resize((ancho, alto), Image.LANCZOS)
                calidad = calidad_inicial 

            if ancho < 200 or alto < 200:
                print(f"   → Comprimida (mín. tamaño): {tamano_kb:.2f} KB")
                return

    except Exception as e:
        print(f"   Error al procesar {ruta}: {e}")

# --- FASE 1: Escaneo y Conteo ---
print("Iniciando escaneo recursivo para conteo...")

# 1. Obtener archivos de destino (lo que ya existe)
# Usamos 'set' para que la búsqueda sea instantánea
try:
    archivos_destino_existentes = set(os.listdir(carpeta_salida))
    print(f"Encontrados {len(archivos_destino_existentes)} archivos en el destino.")
except Exception as e:
    print(f"Error leyendo carpeta de salida: {e}")
    archivos_destino_existentes = set()

# 2. Recorrer el origen y comparar
pendientes = [] # Almacena tuplas (ruta_origen, nombre_destino)
total_encontradas = 0
total_procesadas_previamente = 0

for root, dirs, files in os.walk(carpeta):
    for file in files:
        # Solo procesar archivos de imagen
        if not file.lower().endswith(('.jpg', '.jpeg', '.png')):
            continue
        
        total_encontradas += 1
        ruta_origen = os.path.join(root, file)
        
        # Calcular el nombre de destino final (manejando .png -> .jpg)
        basename, ext = os.path.splitext(file)
        if ext.lower() == '.png':
            destino_name = basename + ".jpg"
        else:
            destino_name = file
        
        # Decidir si está pendiente
        if destino_name in archivos_destino_existentes:
            total_procesadas_previamente += 1
        else:
            pendientes.append((ruta_origen, destino_name))

total_pendientes = len(pendientes)

# --- FASE 2: Resumen Inicial ---
print("\n--- Resumen del Escaneo ---")
print(f"Imágenes totales encontradas: {total_encontradas}")
print(f"Imágenes ya procesadas (saltadas): {total_procesadas_previamente}")
print(f"Imágenes pendientes por procesar: {total_pendientes}")
print("--------------------------------\n")

# --- FASE 3: Bucle de Procesamiento con Porcentaje ---
if total_pendientes == 0:
    print("¡No hay imágenes pendientes! Todo está al día.")
else:
    print(f"Iniciando procesamiento de {total_pendientes} imágenes...")
    
    for i, (ruta_origen, destino_name) in enumerate(pendientes):
        # Calcular porcentaje
        porcentaje = ((i + 1) / total_pendientes) * 100
        
        # Construir la ruta de destino completa
        ruta_destino = os.path.join(carpeta_salida, destino_name)
        
        try:
            tamano_kb = os.path.getsize(ruta_origen) / 1024
            # Imprimir con el nuevo formato de porcentaje
            print(f"[{i+1}/{total_pendientes} | {porcentaje:.1f}%] {os.path.relpath(ruta_origen, carpeta)} ({tamano_kb:.2f} KB)")
            
            # Llamada a la función
            procesar_imagen(ruta_origen, ruta_destino, limite_kb)
            
        except Exception as e:
            print(f"   Error al procesar {ruta_origen}: {e}")

print("\n--- Proceso Completado ---")