from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import os
from datetime import datetime, timedelta    

# Obtener la fecha del día anterior
fecha_ayer = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")

# Ruta de la carpeta con imágenes
carpeta_imagenes = r"C:\Users\jachamizo\Pictures\Reporte de ventas\2025-04-21"  # Ruta absoluta
#carpeta_imagenes = r"C:\Users\jachamizo\Pictures\Reporte de ventas\prueba"  # Ruta absoluta de pruebas 
# Nombre del grupo al que se enviarán las imágenes
#nombre_grupo = "My self"
nombre_grupo = "Supervisores"
# Conectar a Chrome en modo depuración (asegúrate de que esté abierto con --remote-debugging-port=9222)
chrome_options = webdriver.ChromeOptions()
chrome_options.debugger_address = "localhost:9222"  # Se conecta a la sesión abierta

# Iniciar WebDriver con la sesión existente
driver = webdriver.Chrome(options=chrome_options)

# Abrir WhatsApp Web
driver.get("https://web.whatsapp.com/")
time.sleep(7)  # Esperar a que cargue

# Buscar el grupo de WhatsApp
search_box = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
search_box.send_keys(nombre_grupo)
time.sleep(3)  # Esperar que aparezcan los resultados
search_box.send_keys(Keys.ENTER)
print(f"Grupo encontrado: {nombre_grupo}")

# Enviar mensaje inicial
mensaje_inicial = f"se comparte el ranking de ventas por equipo hasta el cierre de ayer {fecha_ayer}"
#mensaje_inicial = f"Buenos días a todos, se comparte el ranking de ventas por equipo hasta el cierre de este fin de semana"
message_box = driver.find_element(By.XPATH, "//div[@aria-label='Type a message']")
message_box.send_keys(mensaje_inicial)
time.sleep(1)
message_box.send_keys(Keys.ENTER)
time.sleep(3)
print("Mensaje inicial enviado.")

# Recorrer imágenes en la carpeta y enviarlas
for imagen in os.listdir(carpeta_imagenes):
    if imagen.endswith((".png", ".jpg", ".jpeg")):
        try:
            # Obtener el nombre de la imagen sin la extensión
            nombre_sin_extension = os.path.splitext(imagen)[0]
            # remplazar los guiones por espacios SUPERVISOR__DENIS_BISBAL -> DENIS BISBAL y quitar la parte de SUPERVISOR 
            nombre_sin_extension = nombre_sin_extension.replace("__", " ").replace("SUPERVISOR ", "").strip()
            nombre_sin_extension = nombre_sin_extension.replace("_", " ").replace("SUPERVISOR ", "").strip()
            print(f"Procesando imagen: {nombre_sin_extension}")
            
            # Hacer clic en el icono de adjuntar
            attach_btn = driver.find_element(By.XPATH, "//span[@data-icon='plus']")
            attach_btn.click()
            time.sleep(1)

            # Subir la imagen
            img_path = os.path.abspath(os.path.join(carpeta_imagenes, imagen))
            upload_input = driver.find_element(By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")
            upload_input.send_keys(img_path)
            time.sleep(3)

            # Agregar descripción con el nombre de la imagen sin extensión
            caption_box = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
            caption_box.send_keys(f"Avance de ventas del equipo de {nombre_sin_extension}")
            time.sleep(1)

            # Enviar la imagen
            send_button = driver.find_element(By.XPATH, "//span[@data-icon='send']")
            send_button.click()
            time.sleep(3)

            print(f"Imagen enviada: {imagen}")

        except Exception as e:
            print(f"Error al enviar {imagen}: {e}")

print("Todas las imágenes fueron procesadas.")
driver.quit()
