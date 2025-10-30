import os
import shutil
from PIL import Image

# --- NO HAY CAMBIOS AQUÍ ---
# Carpetas
carpeta = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TD GUAFITA\2.INFOENTRADA\DATOS_CRUDOS\Eliana\29 OCT"
carpeta_salida = r"\\192.168.1.2\cenyt-proyectos\CEN-223_TOP DRIVE\2.INFOENTRADA\DATOS_FINALES/FOTOS_TRATADAS"
os.makedirs(carpeta_salida, exist_ok=True)
# Límite en KB
limite_kb = 500

# Archivos pendientes
archivos_entrada = {f for f in os.listdir(carpeta) if f.lower().endswith(('.jpg', '.jpeg', '.png'))}
archivos_salida = {f for f in os.listdir(carpeta_salida) if f.lower().endswith(('.jpg', '.jpeg', '.png'))}
pendientes = archivos_entrada - archivos_salida

print(f"Total imágenes en carpeta: {len(archivos_entrada)}")
print(f"Ya procesadas: {len(archivos_salida)}")
print(f"Pendientes: {len(pendientes)}")


# --- FUNCIÓN MODIFICADA ---
# Se renombra a 'procesar_imagen' y ahora maneja la rotación y la compresión.
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
        
        # --- ¡NUEVO PASO! ---
        # 1. Verificar dimensiones y rotar si es necesario
        ancho, alto = img.size
        fue_rotada = False
        if ancho > alto:
            # Rotar 90 grados a la derecha
            # PIL.Image.ROTATE_270 es 90 grados en sentido horario
            img = img.transpose(Image.ROTATE_270)
            fue_rotada = True
            # Actualizamos las dimensiones después de rotar
            ancho, alto = img.size 

        # 2. Manejar PNG -> JPG (lógica que ya tenías)
        if ruta.lower().endswith(".png"):
            img = img.convert("RGB")
            # Actualizar el destino para que sea .jpg
            destino = os.path.splitext(destino)[0] + ".jpg"

        # 3. Decidir si comprimir o solo guardar
        
        # Caso A: Estaba por debajo del límite Y NO necesitaba cambios (ni rotar, ni PNG)
        if tamano_original_kb <= limite_kb and not fue_rotada and not ruta.lower().endswith(".png"):
            shutil.copy2(ruta, destino)
            print(f"  → Copiada sin cambios ({tamano_original_kb:.2f} KB)")
            return

        # Caso B: Estaba por debajo del límite PERO fue rotada o era PNG
        if tamano_original_kb <= limite_kb:
            img.save(destino, quality=95, optimize=True) # Guardar con alta calidad
            nuevo_tamano = os.path.getsize(destino) / 1024
            print(f"  → Guardada (Rotada/Convertida): {nuevo_tamano:.2f} KB")
            return

        # --- Lógica de compresión (casi igual a tu función original) ---
        # Caso C: Superaba el límite (necesita compresión)
        calidad = calidad_inicial

        while True:
            # Guardar temporalmente
            img.save(destino, optimize=True, quality=calidad)
            tamano_kb = os.path.getsize(destino) / 1024

            if tamano_kb <= limite_kb:
                print(f"  → Comprimida: {tamano_kb:.2f} KB")
                return # ¡Logrado!

            # Bajar calidad primero
            if calidad > 20:
                calidad -= paso
            else:
                # Si la calidad ya está muy baja, reducimos dimensiones
                ancho = int(ancho * factor_resize)
                alto = int(alto * factor_resize)
                img = img.resize((ancho, alto), Image.LANCZOS)
                # Al reducir tamaño, reseteamos la calidad para intentarlo de nuevo
                calidad = calidad_inicial 

            # Si llega a ser muy pequeña y no logra el peso, salimos
            if ancho < 200 or alto < 200:
                print(f"  → Comprimida (mín. tamaño): {tamano_kb:.2f} KB")
                return # Salir aunque no cumpla el límite

    except Exception as e:
        print(f"  Error al procesar {ruta}: {e}")

# --- BUCLE PRINCIPAL MODIFICADO ---
# Ahora solo llama a la nueva función 'procesar_imagen'
for archivo in pendientes:
    ruta = os.path.join(carpeta, archivo)
    destino = os.path.join(carpeta_salida, archivo)

    if os.path.isfile(ruta):
        tamano_kb = os.path.getsize(ruta) / 1024
        print(f"{archivo} -> {tamano_kb:.2f} KB")
        
        # Llamada única a la nueva función que maneja toda la lógica
        procesar_imagen(ruta, destino, limite_kb)