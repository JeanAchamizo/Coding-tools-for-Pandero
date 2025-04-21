import time
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
import ast

class ContratoAutomatizador:
    def __init__(self, google_sheet_id, columna_contrato=5, columna_cuota=14, columna_estado=16, chrome_debugger_address="localhost:9222"):
        self.google_sheet_id = google_sheet_id
        self.columna_contrato = columna_contrato
        self.columna_cuota = columna_cuota
        self.columna_estado = columna_estado
        self.chrome_debugger_address = chrome_debugger_address
        self.driver = self._iniciar_driver()

    def _iniciar_driver(self):
        chrome_options = webdriver.ChromeOptions()
        chrome_options.debugger_address = self.chrome_debugger_address
        return webdriver.Chrome(options=chrome_options)

    @staticmethod
    def letra_a_numero_columna(letra_columna):
        numero_columna = 0
        for letra in letra_columna:
            numero_columna = numero_columna * 26 + (ord(letra.upper()) - ord('A') + 1)
        return numero_columna

    def obtener_dato_google_sheet(self, row, column):
        url = f"https://script.google.com/macros/s/{self.google_sheet_id}/exec?row={row}&column={column}"
        response = requests.get(url)
        contrato = response.text
        return contrato.split(":")[1].strip() if contrato else None #7507-080-26 

    def obtener_datos_google_sheet(self, startRow, endRow, column):
        url = f"https://script.google.com/macros/s/{self.google_sheet_id}/exec?action=getRange&startRow={startRow}&endRow={endRow}&column={column}"
        response = requests.get(url)
        
        if response.status_code == 200:
            try:
                # La respuesta es un string como "response: [6695-088-12,4343-028-35,4349-065-20,7509-028-13,4339-063-42]"
                # Eliminar el prefijo "response: " y procesar el contenido
                datos = response.text.replace("response: [", "").replace("]", "").strip()
                # Dividir los valores por coma y convertirlos en una lista
                lista_datos = [dato.strip() for dato in datos.split(",")]
                return lista_datos
            except Exception as e:
                print(f"Error al procesar la respuesta: {e}")
                return []
        else:
            print(f"Error al obtener datos: {response.status_code} - {response.text}")
            return []
    
    def actualizar_google_sheet(self, row, column, value):
        url = f"https://script.google.com/macros/s/{self.google_sheet_id}/exec?row={row}&column={column}&newValue={value}"
        response = requests.post(url)
        print(f"Actualización en Sheets en fila {row}: {response.text}")

    def actualizar_rango_google_sheet(self, start_row, end_row, column, values):
            # Convertir la lista de valores en un formato JSON válido
            values_json = str(values).replace("'", '"')  # Convertir comillas simples a dobles
            url = f"https://script.google.com/macros/s/{self.google_sheet_id}/exec?action=updateRange&startRow={start_row}&endRow={end_row}&column={column}&values={values_json}"
            response = requests.post(url)
            if response.status_code == 200:
                print(f"Actualización exitosa para el rango {start_row}-{end_row} en la columna {column}.")
            else:
                print(f"Error al actualizar el rango: {response.status_code} - {response.text}")

    def verificar_contrato_en_intranet(self, contrato):
        # Abrir la página de la intranet
        self.driver.get("https://saf.pandero.com.pe/saf/reportes/ReporteEstadoCuenta.aspx")

        # Seleccionar la opción de búsqueda
        attach_btn = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_maincontent_BsqContrato_ddlTipoBusqueda'))
        )
        selecttag = Select(attach_btn)
        selecttag.select_by_value('1')

        # Ingresar el número de contrato
        campo_contrato = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, 'ctl00_maincontent_BsqContrato_txtNumeroDocumento'))
        )
        print("Escribiendo contrato:", contrato)
        self.driver.execute_script(
            "arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));",
            campo_contrato, contrato
        )

        # Primer clic en el botón de búsqueda
        boton_buscar = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_maincontent_BsqContrato_btnBuscar'))
        )
        boton_buscar.click()

        # Esperar el mensaje "No hay registros que mostrar"
        try:
            WebDriverWait(self.driver, 5).until(
                EC.text_to_be_present_in_element(
                    (By.XPATH, "//table[@id='ctl00_maincontent_BsqContrato_gv_Contratos']//center"),
                    "No hay registros que mostrar"
                )
            )
        except TimeoutException:
            print("No se detectó el mensaje 'No hay registros que mostrar'. Continuando...")

        # Segundo clic en el botón de búsqueda
        boton_buscar = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_maincontent_BsqContrato_btnBuscar'))
        )
        boton_buscar.click()

        # Verificar si hay resultados reales en la tabla
        try:
            WebDriverWait(self.driver, 5).until_not(
                EC.text_to_be_present_in_element(
                    (By.XPATH, "//table[@id='ctl00_maincontent_BsqContrato_gv_Contratos']//center"),
                    "No hay registros que mostrar"
                )
            )

            # Hacer clic en el botón "Seleccionar"
            boton_selec = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, 'ctl00_maincontent_BsqContrato_btnSeleccionar'))
            )
            boton_selec.click()
        except TimeoutException:
            return 0

        # Obtener cuotas pagadas y por pagar
        couta_pagada = float(WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_maincontent_lblPagadoA'))
        ).text.replace(",", ""))  # Eliminar las comas antes de convertir a float
        
        couta_por_pagar = float(WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ctl00_maincontent_lblPorPagarA'))
        ).text.replace(",", ""))  # Eliminar las comas antes de convertir a float
        
        print("Cuota pagada:", couta_pagada)
        print("Cuota por pagar:", couta_por_pagar)

        # Buscar la última fecha de pago para "CIA"
        ultima_fecha_cia = self.obtener_ultima_fecha_pago_cia()

        # Verificar si la cuota está completamente pagada
        if couta_pagada > 0 and couta_por_pagar == 0:
            return couta_pagada, ultima_fecha_cia
        return 0, ultima_fecha_cia

    def obtener_ultima_fecha_pago_cia(self):
        try:
            # Localizar la tabla de pagos
            tabla_pagos = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'ctl00_maincontent_griPagos'))
            )
            
            # Obtener todas las filas de la tabla
            filas = tabla_pagos.find_elements(By.XPATH, ".//tr[@class='row']")
            ultima_fecha = None
    
            for fila in filas:
                # Obtener la columna que contiene la descripción (donde puede aparecer "CIA")
                descripcion = fila.find_element(By.XPATH, "./td[13]").text.strip()  # Columna 13
                if "CIA" in descripcion:
                    # Si contiene "CIA", obtener la fecha de la primera columna
                    fecha = fila.find_element(By.XPATH, "./td[1]").text.strip()
                    ultima_fecha = fecha  # Actualizar la última fecha encontrada
    
            print(f"Última fecha de pago para 'CIA': {ultima_fecha}")
            return ultima_fecha
        except Exception as e:
            print(f"Error al obtener la última fecha de pago para 'CIA': {e}")
            return None

    def automatizar_proceso(self, row_inicial=2, row_final=2):
        # Obtener los contratos en el rango especificado
        contratos = self.obtener_datos_google_sheet(row_inicial, row_final, self.columna_contrato)
        cuotas_pagadas = []  # Lista para almacenar las cuotas pagadas
        estados = []  # Lista para almacenar los estados (Pagado/No Pagado)
    
        for contrato in contratos:
            print("Procesando contrato:", contrato)
            cuota_pagada, ultima_fecha_cia = self.verificar_contrato_en_intranet(contrato)
            cuotas_pagadas.append(cuota_pagada)
            estados.append(ultima_fecha_cia if ultima_fecha_cia else "No Pagado")

        # Enviar las cuotas pagadas al endpoint en un solo llamado
        print("Actualizando cuotas pagadas en Google Sheets...")
        self.actualizar_rango_google_sheet(row_inicial, row_final, self.columna_cuota, cuotas_pagadas)
    
        # Enviar los estados al endpoint en un solo llamado
        self.actualizar_rango_google_sheet(row_inicial, row_final, self.columna_estado, estados)

# Ejecutar la automatización
if __name__ == "__main__":

    #Link de prueba
    google_sheet_id = "AKfycbz2RL-OD9quDrOT_MRqFkNrnttTRIxZHWSwwByv7WgR2m5SV_KVWxtsbitTV_drwJ9Y"
    #Link de producción
    #google_sheet_id = "AKfycbz9vkiLAOm5VucnSadIIdUQbWUzoCgNAHogtxKMSUFkN9cmyj-I1qNxQg5QGL1IQ78QRw" # ID de tu Google Sheet
    
    # Convertir letras de columna a números // cambia la columna según tu hoja de cálculo
    columna_contrato = ContratoAutomatizador.letra_a_numero_columna("E")  # Salida: 5
    columna_cuota = ContratoAutomatizador.letra_a_numero_columna("L")     # Salida: 14
    columna_estado = ContratoAutomatizador.letra_a_numero_columna("N")    # Salida: 16

    fila_inicial = 2 # Fila inicial para la automatización
    fila_final = 17 # Fila final para la automatización

    automatizador = ContratoAutomatizador(google_sheet_id, columna_contrato, columna_cuota, columna_estado)
    automatizador.automatizar_proceso(fila_inicial,fila_final)