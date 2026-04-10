import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv


# --- CONFIGURACIÓN ---
USUARIO = os.getenv("CH_USER")
CONTRASEÑA = os.getenv("CH_PASS")
URL_LOGIN = os.getenv("URL_SISTEMA")
URL_REPORTES = os.getenv("CH_URL")

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

try:
    # 1. Login
    driver.get(URL_LOGIN)
    time.sleep(2)
    driver.find_element(By.ID, "user_login").send_keys(USUARIO)
    driver.find_element(By.ID, "user_pass").send_keys(CONTRASEÑA)
    driver.find_element(By.ID, "wp-submit").click()
    time.sleep(3)

    # 2. Ir a la página de reportes
    driver.get(URL_REPORTES)
    time.sleep(4)

    # 3. Extraer nombres del selector
    select_element = driver.find_element(By.ID, "reporte_gerencial_empresa")
    options_elements = select_element.find_elements(By.TAG_NAME, "option")

    lista_empresas = [opt.text for opt in options_elements if opt.text != "Seleccione una empresa" and opt.text.strip() != ""]

    print("\n--- LISTA DE EMPRESAS ENCONTRADAS (Copia desde aquí) ---")
    print("LISTA_GENERAL = [")
    for e in lista_empresas:
        print(f'    "{e}",')
    print("]")
    print("------------------------------------------------------")

finally:
    driver.quit()