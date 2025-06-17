import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QStackedWidget, QMenuBar, QAction, QFileDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from Contratos import ContratoAutomatizador
import subprocess


# --- Worker para Verificar CIA ---
class WorkerVerificarCIA(QThread):
    output_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, google_sheet_id, col_contrato, col_cuota, col_estado, fila_ini, fila_fin):
        super().__init__()
        self.google_sheet_id = google_sheet_id
        self.col_contrato = col_contrato
        self.col_cuota = col_cuota
        self.col_estado = col_estado
        self.fila_ini = fila_ini
        self.fila_fin = fila_fin

    def run(self):
        import builtins
        original_print = print
        def custom_print(*args, **kwargs):
            msg = ' '.join(str(a) for a in args)
            self.output_signal.emit(msg)
            original_print(*args, **kwargs)
        builtins.print = custom_print

        try:
            automatizador = ContratoAutomatizador(
                self.google_sheet_id,
                self.col_contrato,
                self.col_cuota,
                self.col_estado
            )
            automatizador.automatizar_proceso(self.fila_ini, self.fila_fin)
        except Exception as e:
            self.output_signal.emit(f"Error: {e}")
        finally:
            builtins.print = original_print
            self.finished_signal.emit()

# --- Worker para Enviar Reporte WhatsApp ---
class WorkerEnviarReporte(QThread):
    output_signal = pyqtSignal(str)
    finished_signal = pyqtSignal()

    def __init__(self, mensaje, carpeta, grupo):
        super().__init__()
        self.mensaje = mensaje
        self.carpeta = carpeta
        self.grupo = grupo

    def run(self):
        import builtins
        original_print = print
        def custom_print(*args, **kwargs):
            msg = ' '.join(str(a) for a in args)
            self.output_signal.emit(msg)
            original_print(*args, **kwargs)
        builtins.print = custom_print

        try:
            from selenium import webdriver
            from selenium.webdriver.common.by import By
            from selenium.webdriver.common.keys import Keys
            import time
            import os
            from datetime import datetime, timedelta

            fecha_ayer = (datetime.now() - timedelta(days=1)).strftime("%d-%m-%Y")
            carpeta_imagenes = self.carpeta
            nombre_grupo = self.grupo
            mensaje_inicial = self.mensaje.replace("{fecha}", fecha_ayer)

            chrome_options = webdriver.ChromeOptions()
            chrome_options.debugger_address = "localhost:9222"
            driver = webdriver.Chrome(options=chrome_options)
            driver.get("https://web.whatsapp.com/")
            time.sleep(7)

            search_box = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
            search_box.send_keys(nombre_grupo)
            time.sleep(3)
            search_box.send_keys(Keys.ENTER)
            print(f"Grupo encontrado: {nombre_grupo}")

            message_box = driver.find_element(By.XPATH, "//div[@aria-label='Type a message']")
            message_box.send_keys(mensaje_inicial)
            time.sleep(1)
            message_box.send_keys(Keys.ENTER)
            time.sleep(3)
            print("Mensaje inicial enviado.")

            for imagen in os.listdir(carpeta_imagenes):
                if imagen.endswith((".png", ".jpg", ".jpeg")):
                    try:
                        nombre_sin_extension = os.path.splitext(imagen)[0]
                        nombre_sin_extension = nombre_sin_extension.replace("__", " ").replace("SUPERVISOR ", "").strip()
                        nombre_sin_extension = nombre_sin_extension.replace("_", " ").replace("SUPERVISOR ", "").strip()
                        print(f"Procesando imagen: {nombre_sin_extension}")

                        # Busca cualquier botón de adjuntar con data-icon que contenga "plus"
                        attach_btn = driver.find_element(By.XPATH, "//span[contains(@data-icon, 'plus')]")
                        attach_btn.click()
                        time.sleep(1)

                        img_path = os.path.abspath(os.path.join(carpeta_imagenes, imagen))
                        upload_input = driver.find_element(By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")
                        upload_input.send_keys(img_path)
                        time.sleep(3)

                        caption_box = driver.find_element(By.XPATH, "//div[@contenteditable='true']")
                        caption_box.send_keys(f"Avance de ventas del equipo de {nombre_sin_extension}")
                        time.sleep(1)

                        # Busca cualquier botón de enviar con data-icon que contenga "send"
                        send_button = driver.find_element(By.XPATH, "//span[contains(@data-icon, 'send')]")
                        send_button.click()
                        time.sleep(3)

                        print(f"Imagen enviada: {imagen}")

                    except Exception as e:
                        print(f"Error al enviar {imagen}: {e}")

            print("Todas las imágenes fueron procesadas.")
            driver.quit()
        except Exception as e:
            self.output_signal.emit(f"Error: {e}")
        finally:
            builtins.print = original_print
            self.finished_signal.emit()

# --- Ventana Principal ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Automatización Pandero - Menú")
        self.setGeometry(100, 100, 700, 500)

        self.stacked = QStackedWidget()
        self.setCentralWidget(self.stacked)

        # Menú
        menubar = QMenuBar(self)
        self.setMenuBar(menubar)
        menu = menubar.addMenu("Opciones")
        action_verificar = QAction("Verificar CIA", self)
        action_reporte = QAction("Enviar Reporte WhatsApp", self)
        menu.addAction(action_verificar)
        menu.addAction(action_reporte)
        action_verificar.triggered.connect(lambda: self.stacked.setCurrentIndex(0))
        action_reporte.triggered.connect(lambda: self.stacked.setCurrentIndex(1))

        # Botón para abrir Chrome en modo debugging (al lado del menú)
        self.btn_abrir_chrome = QPushButton("Abrir Chrome en modo debugging", self)
        self.btn_abrir_chrome.clicked.connect(self.abrir_chrome_debug)
        menubar.setCornerWidget(self.btn_abrir_chrome, Qt.TopRightCorner)


        # Página Verificar CIA
        self.page_verificar = QWidget()
        layout1 = QVBoxLayout()
        self.input_sheet = QLineEdit()
        self.input_sheet.setPlaceholderText("ID de Google Sheet")
        self.input_sheet.setText("AKfycbz2RL-OD9quDrOT_MRqFkNrnttTRIxZHWSwwByv7WgR2m5SV_KVWxtsbitTV_drwJ9Y")
        layout1.addWidget(QLabel("ID de Google Sheet:"))
        layout1.addWidget(self.input_sheet)
        col_layout = QHBoxLayout()
        self.input_col_contrato = QLineEdit("E")
        self.input_col_cuota = QLineEdit("L")
        self.input_col_estado = QLineEdit("N")
        col_layout.addWidget(QLabel("Contrato:"))
        col_layout.addWidget(self.input_col_contrato)
        col_layout.addWidget(QLabel("Cuota:"))
        col_layout.addWidget(self.input_col_cuota)
        col_layout.addWidget(QLabel("Estado:"))
        col_layout.addWidget(self.input_col_estado)
        layout1.addLayout(col_layout)
        fila_layout = QHBoxLayout()
        self.input_fila_ini = QLineEdit("2")
        self.input_fila_fin = QLineEdit("20")
        fila_layout.addWidget(QLabel("Fila inicial:"))
        fila_layout.addWidget(self.input_fila_ini)
        fila_layout.addWidget(QLabel("Fila final:"))
        fila_layout.addWidget(self.input_fila_fin)
        layout1.addLayout(fila_layout)
        self.btn_ejecutar1 = QPushButton("Ejecutar Verificación CIA")
        self.btn_ejecutar1.clicked.connect(self.ejecutar_verificar_cia)
        layout1.addWidget(self.btn_ejecutar1)
        self.output1 = QTextEdit()
        self.output1.setReadOnly(True)
        layout1.addWidget(QLabel("Salida de la consola:"))
        layout1.addWidget(self.output1)
        self.page_verificar.setLayout(layout1)

        # Página Enviar Reporte WhatsApp
        self.page_reporte = QWidget()
        layout2 = QVBoxLayout()
        self.input_mensaje = QLineEdit("Buenos días a todos, se comparte el ranking de ventas por equipo hasta el cierre de ayer {fecha}")
        layout2.addWidget(QLabel("Mensaje inicial (usa {fecha} para la fecha de ayer):"))
        layout2.addWidget(self.input_mensaje)
        carpeta_layout = QHBoxLayout()
        self.input_carpeta = QLineEdit(r"C:\Users\jachamizo\Pictures\Reporte de ventas\2025-06-1")
        btn_carpeta = QPushButton("Seleccionar carpeta")
        btn_carpeta.clicked.connect(self.seleccionar_carpeta)
        carpeta_layout.addWidget(QLabel("Carpeta de imágenes:"))
        carpeta_layout.addWidget(self.input_carpeta)
        carpeta_layout.addWidget(btn_carpeta)
        layout2.addLayout(carpeta_layout)
        self.input_grupo = QLineEdit("Supervisores")
        layout2.addWidget(QLabel("Nombre del grupo de WhatsApp:"))
        layout2.addWidget(self.input_grupo)
        self.btn_ejecutar2 = QPushButton("Enviar Reporte por WhatsApp")
        self.btn_ejecutar2.clicked.connect(self.ejecutar_enviar_reporte)
        layout2.addWidget(self.btn_ejecutar2)
        self.output2 = QTextEdit()
        self.output2.setReadOnly(True)
        layout2.addWidget(QLabel("Salida de la consola:"))
        layout2.addWidget(self.output2)
        self.page_reporte.setLayout(layout2)

        # Añadir páginas al stacked widget
        self.stacked.addWidget(self.page_verificar)
        self.stacked.addWidget(self.page_reporte)

    def abrir_chrome_debug(self):
        chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
        user_data_dir = r"C:\selenium\ChromeProfile"
        comando = [
            chrome_path,
            "--remote-debugging-port=9222",
            f'--user-data-dir={user_data_dir}'
        ]
        try:
            subprocess.Popen(comando)
            if hasattr(self, "output1"):
                self.output1.append("Chrome en modo debugging iniciado.")
            if hasattr(self, "output2"):
                self.output2.append("Chrome en modo debugging iniciado.")
        except Exception as e:
            if hasattr(self, "output1"):
                self.output1.append(f"Error al abrir Chrome: {e}")
            if hasattr(self, "output2"):
                self.output2.append(f"Error al abrir Chrome: {e}")

    def seleccionar_carpeta(self):
        carpeta = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de imágenes")
        if carpeta:
            self.input_carpeta.setText(carpeta)

    def ejecutar_verificar_cia(self):
        self.output1.clear()
        google_sheet_id = self.input_sheet.text().strip()
        col_contrato = ContratoAutomatizador.letra_a_numero_columna(self.input_col_contrato.text().strip())
        col_cuota = ContratoAutomatizador.letra_a_numero_columna(self.input_col_cuota.text().strip())
        col_estado = ContratoAutomatizador.letra_a_numero_columna(self.input_col_estado.text().strip())
        try:
            fila_ini = int(self.input_fila_ini.text().strip())
            fila_fin = int(self.input_fila_fin.text().strip())
        except ValueError:
            self.output1.append("Error: Las filas deben ser números enteros.")
            return

        self.btn_ejecutar1.setEnabled(False)
        self.worker1 = WorkerVerificarCIA(
            google_sheet_id, col_contrato, col_cuota, col_estado, fila_ini, fila_fin
        )
        self.worker1.output_signal.connect(self.output1.append)
        self.worker1.finished_signal.connect(lambda: self.btn_ejecutar1.setEnabled(True))
        self.worker1.start()

    def ejecutar_enviar_reporte(self):
        self.output2.clear()
        mensaje = self.input_mensaje.text().strip()
        carpeta = self.input_carpeta.text().strip()
        grupo = self.input_grupo.text().strip()
        if not os.path.isdir(carpeta):
            self.output2.append("Error: La carpeta de imágenes no existe.")
            return
        if not grupo:
            self.output2.append("Error: Debes ingresar el nombre del grupo de WhatsApp.")
            return

        self.btn_ejecutar2.setEnabled(False)
        self.worker2 = WorkerEnviarReporte(mensaje, carpeta, grupo)
        self.worker2.output_signal.connect(self.output2.append)
        self.worker2.finished_signal.connect(lambda: self.btn_ejecutar2.setEnabled(True))
        self.worker2.start()

if __name__ == "__main__":
    import os
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())