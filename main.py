# -------- SETUP --------

# Importación de Selenium y Pandas.
from selenium import webdriver 
import pandas as pd

# Importaciones extra de Selenium.
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.chrome.options import Options

# Importaciones de Email.
from email.mime.application import  MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


# Importaciones de Google Drive
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
# from pydrive.auth import GoogleAuth
# from pydrive.drive import GoogleDrive
import os.path

googleCredentials = "client_secret_613186623967-9bjccboujjn3mg5sskeej0tbumi30up3.apps.googleusercontent.com.json"
SCOPES = ['https://www.googleapis.com/auth/drive.file']
API_SERVICE_NAME = 'drive'
API_VERSION = 'v3'

# Otras importaciones y declaraciones.
from datetime import datetime
import logging
import json
import smtplib
import os

url_banco = "https://www.bna.com.ar/Personas"
url_dolarblue = "https://dolarhoy.com"
table_element = (By.CLASS_NAME, "cotizacion")
today = datetime.now().strftime("%d/%m/%Y")
csv_filename = "BCA Compra-Venta.csv"
email_CC = "example@gmail.com"
log_file_path = "logs/bot_log.txt"

# Configuración del logging

logging.basicConfig(filename=log_file_path, level=logging.INFO, format='%(asctime)s - %(levelname)s: %(message)s')

# Limpiar el archivo existente.
if os.path.exists(log_file_path):
    logging.shutdown()
    os.remove(log_file_path)

logging.info("Importacion de librerías y configuración del logging correcta. Inicializando proceso:")

# -------- FUNCIONES --------
# Configuración de credenciales Google Drive.
def get_credentials():
    try:
        logging.info("Inicializando get_credentials.")
        creds = None

        credentialsPath = os.path.join("credentials", googleCredentials)

        if(os.path.exists('token.json')):
            creds = Credentials.from_authorized_user_file('token.json')

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(googleCredentials, SCOPES)
                creds = flow.run_local_server(port=0)

        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
        return creds
    except Exception as e:
        logging.error(f"Ocurrió el siguiente error durante la obtención de las credenciales de Google: {str(e)}")
        raise


# Inicialización del driver para la página del banco. 
def init_driver():
    try: 
        logging.info("Inicializando driver del Banco Nación con Google Chrome.")
        # Configuración de Google Chrome.
        chrome_options = Options()
        driver = webdriver.Chrome(options=chrome_options)
        driver.maximize_window()

        # Navegación hasta la página en cuestión.
        driver.get(url_banco)

        # Localizar elemento de cotizaciones una vez esté cargado.
        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.TAG_NAME, "body")))
        WebDriverWait(driver, 10).until(ec.presence_of_element_located(table_element))

        logging.info("Inicialización del driver completada.")
        return driver
    
    except Exception as e:
        logging.error(f"Ocurrió el siguiente error durante el inicio del navegador: {str(e)}")
        raise

# Utiliza el driver para devolver la información de la tabla cotizaciones.
def get_currency_data(driver):
    try:
        logging.info("Obteniendo información en la página del Banco Nación.")
        # Obtención de información de la tabla cotizaciones.
        table = driver.find_element(*table_element)
        rows = table.find_elements(By.TAG_NAME, "tr")

        # Extración de datos.
        data = []
        for i, row in enumerate(rows):
            if i > 0:
                columns = row.find_elements(By.TAG_NAME, "td")
                data.append([col.text.strip() for col in columns])
        logging.info("Información del Banco obtenida correctamente.")
        return data
    except Exception as e:
        logging.error(f"Ocurrió el siguiente error durante la obtención de datos de la tabla: {str(e)}")
        raise

# Creación y utlización de otro Chrome driver para obtener información actualizada del valor del Dolar Blue.
def get_dolarblue_data():
    try:
        logging.info("Inicializando driver y obteniendo información del Dólar Blue.")
        # Configuración de Google Chrome.
        chrome_options = Options()
        driver = webdriver.Chrome(options=chrome_options)
        driver.maximize_window()

        # Navegación hasta DolarHoy.
        driver.get(url_dolarblue)

        # Esperar a que cargue.
        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.TAG_NAME, "body")))

        # Encontrar el elemento "values", donde se encuentran lógicamente los valores.
        elementoValues = driver.find_element(By.CLASS_NAME, "values")
        elementoCompra = elementoValues.find_element(By.CLASS_NAME, "compra")
        elementoVenta = elementoValues.find_element(By.CLASS_NAME, "venta")

        valorCompra = elementoCompra.find_element(By.CLASS_NAME, "val").text.replace("$", "")
        valorVenta = elementoVenta.find_element(By.CLASS_NAME, "val").text.replace("$", "")

        driver.quit()
        logging.info("Información del Dólar Blue obtenida correctamente.")
        return valorCompra, valorVenta
    except Exception as e:
        logging.error(f"Ocurrió el siguiente error durante la obtención del valor Dolar Blue: {str(e)}")
    raise

# Formatizar la información para posteriormente guardarla como CSV.
def format_and_save(data):
    try:
        logging.info("Inicializando format_and_save.")
        # Guardar información en un dataFrame.
        dataFrame = pd.DataFrame(data, columns=["Moneda", "Compra", "Venta"])

        # Agregar fecha
        dataFrame["Fecha"] = today

        # Alinear formato con los valores que luego voy a insertar del Dólar Blue.
        dataFrame[["Compra", "Venta"]] = dataFrame[["Compra", "Venta"]].apply(lambda x: x.str.replace(',', '.').str.replace('"', ''))
        
        # Dejar un solo número decimal después de la coma.
        dataFrame[["Compra", "Venta"]] = dataFrame[["Compra", "Venta"]].apply(lambda x: round(x.astype(float), 1))

        # No mostrar el decimal si el valor es 0.
        dataFrame[["Compra", "Venta"]] = dataFrame[["Compra", "Venta"]].apply(lambda x: x.apply(lambda y: f'{y:g}'))

        # Creat tabla pivot para acomodar que cada fecha tenga todas las monedas.
        dataFrame = dataFrame.pivot(index="Fecha", columns="Moneda")

        # Eliminar niveles adicionales en el índice y reorganizar columnas
        dataFrame.columns = [f"{col[0]} {col[1].lower()}" if col[1] else col[0] for col in dataFrame.columns]

        # Resetear el índice
        dataFrame = dataFrame.reset_index()

        # Obtener y agregar los datos del Dólar Blue
        valorCompra, valorVenta = get_dolarblue_data()
        dataFrame["Venta Blue"] = valorVenta
        dataFrame["Compra Blue"] = valorCompra

        # Renombrar y reorganizar columnas para que queden alineadas con lo que se desea.
        dataFrame.columns = ["Fecha", "Compra Dólar", "Compra Euro", "Compra Real", "Venta Dólar", "Venta Euro", "Venta Real", "Venta Blue", "Compra Blue"]
        columnOrder = ["Fecha", "Venta Dólar", "Venta Blue", "Compra Dólar", "Compra Blue", "Venta Euro", "Compra Euro", "Venta Real", "Compra Real"]
        dataFrame = dataFrame[columnOrder]

        # Si el archivo existe, leerlo y usar lo que ya esté.
        if os.path.exists(csv_filename):
            csv_data = pd.read_csv(csv_filename)
            dataFrame = pd.concat([csv_data, dataFrame], ignore_index=True)

        # Guardar en CSV
        dataFrame.to_csv(csv_filename, index=False)
        logging.info(f"Datos guardados con éxito en el archivo: {csv_filename}")

    except Exception as e:
        logging.error(f"Error durante el guardado de datos en el CSV: {str(e)}")
        raise

def send_email(csv_filename):
    try:
        logging.info("Inicializando envío del csv vía email.")
        # Credenciales de correo electrónico.
        with open('credentials/email_credentials.json') as cred_File:
            credentials = json.load(cred_File)

        # Declaración de variables para el servidor 
        stmp_server = "smtp.gmail.com"
        smtp_port = 587
        sender_email = credentials["sender_email"]
        receiver_email = email_CC

        # Configuración del mensaje.
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Actualización de Cotizaciones"

        # Adjuntar archivo csv al email
        with open(csv_filename, "rb") as file:
            attachment = MIMEApplication(file.read(), _subtype="csv")
            attachment.add_header('Content-Disposition', 'attachment', os.path.basename(csv_filename))
            msg.attach(attachment)

        # Conectar al servidor y enviar el correo
        with smtplib.SMTP(stmp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, credentials["contraseña"])
            server.sendmail(sender_email, receiver_email, msg.as_string())

        logging.info(f"Correo enviado correctamente a {receiver_email}")
    except Exception as e:
        logging.error(f"Error durante el envío del correo: {str(e)}")

if __name__ == "__main__":
    try:
        # Inicio navegador.
        driver = init_driver()

        # Obtención de datos.
        data = get_currency_data(driver)

        # Guardar datos en el csv.
        csv_filename = format_and_save(data)

        # Enviar correo con el adjunto.
        send_email(csv_filename)

        # Cerrar navegador
        driver.quit()

    except Exception as e:
        logging.error(f"Error general: {str(e)}")