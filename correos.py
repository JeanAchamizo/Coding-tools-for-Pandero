import time
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

class ContactosAutomatizador:
    def __init__(self, google_sheet_id, columna_nombres, columna_correos, chrome_debugger_address="localhost:9222"):
        self.google_sheet_id = google_sheet_id
        self.columna_nombres = columna_nombres
        self.columna_correos = columna_correos
        self.chrome_debugger_address = chrome_debugger_address
        self.driver = self._iniciar_driver()

    def _iniciar_driver(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.debugger_address = self.chrome_debugger_address
        return webdriver.Chrome(options=chrome_options)

    def obtener_datos_google_sheet(self, start_row, end_row, column):
        url = f"https://script.google.com/macros/s/{self.google_sheet_id}/exec?action=getRange&startRow={start_row}&endRow={end_row}&column={column}"
        response = requests.get(url)
        
        if response.status_code == 200:
            try:
                # Parsear la respuesta JSON
                datos = response.json().get("response", [])
                
                # Procesar los nombres
                nombres_procesados = []
                for dato in datos:
                    if "," in dato:
                        # Dividir por la coma
                        partes = dato.split(",")
                        antes_de_coma = partes[0].strip()  # Parte antes de la coma
                        despues_de_coma = partes[1].strip().split()[0]  # Primera palabra después de la coma
                        nombres_procesados.append(f"{antes_de_coma} {despues_de_coma}")
                    else:
                        # Si no hay coma, dejar el nombre tal cual
                        nombres_procesados.append(dato.strip())

                return nombres_procesados
            except Exception as e:
                print(f"Error al procesar la respuesta: {e}")
                return []
        else:
            print(f"Error al obtener datos: {response.status_code} - {response.text}")
            return []
    
    def actualizar_google_sheet(self, start_row, end_row, column, values):
        values_json = str(values).replace("'", '"')  # Convertir comillas simples a dobles
        url = f"https://script.google.com/macros/s/{self.google_sheet_id}/exec?action=updateRange&startRow={start_row}&endRow={end_row}&column={column}&values={values_json}"
        response = requests.post(url)
        if response.status_code == 200:
            print(f"Actualización exitosa para el rango {start_row}-{end_row} en la columna {column}.")
        else:
            print(f"Error al actualizar el rango: {response.status_code} - {response.text}")

    def buscar_contactos(self, nombres):
        self.driver.get("https://contacts.google.com/directory")
        time.sleep(5)  # Esperar a que cargue la página

        correos = []  # Lista para almacenar los correos obtenidos

        for nombre in nombres:
            try:
                # Encontrar el input de búsqueda
                input_busqueda = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.Ax4B8.ZAGvjd"))
                )
                input_busqueda.clear()
                input_busqueda.send_keys(nombre)
                time.sleep(2)  # Esperar a que aparezcan los resultados

                # Obtener el resultado de la búsqueda
                resultado = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.MkjOTb.oKubKe"))
                )
                correo = resultado.find_element(By.CSS_SELECTOR, "div.mf6tRb").text.strip()
                
                # Eliminar el carácter especial "‒" si está presente
                correo = correo.replace("‒", "").strip()

                correos.append(correo)
                print(f"Nombre: {nombre}, Correo: {correo}")

                # Limpiar el input para la siguiente búsqueda
                input_busqueda.clear()
                time.sleep(1)
            except TimeoutException:
                print(f"No se encontró correo para: {nombre}")
                correos.append("No encontrado")
        
        return correos

    def automatizar_proceso(self, start_row, end_row):
        # Obtener los nombres desde Google Sheets
        nombres = self.obtener_datos_google_sheet(start_row, end_row, self.columna_nombres)
        # tamaño de nombres
        print(f"numero de Nombres obtenidos: {len(nombres)}")

        # Buscar los correos en Google Contacts
        correos = self.buscar_contactos(nombres)

        # Actualizar los correos en Google Sheets
        self.actualizar_google_sheet(start_row, end_row, self.columna_correos, correos)

# Ejecutar la automatización
if __name__ == "__main__":
    # ID de tu Google Sheet
    google_sheet_id = "AKfycbw8eKuOmjNOo_e08fpq5tQq28WAJ0c9RhxVYcrX_cfarYYwI32xaXPP89ewT4j25mf7"

    # Columnas de nombres y correos
    columna_nombres = 6  # Cambia al número de columna correspondiente
    columna_correos = 8  # Cambia al número de columna correspondiente

    # Rango de filas a procesar
    fila_inicial = 14
    fila_final = 145

    # Crear instancia y ejecutar
    automatizador = ContactosAutomatizador(google_sheet_id, columna_nombres, columna_correos)
    automatizador.automatizar_proceso(fila_inicial, fila_final)