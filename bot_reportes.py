import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from dotenv import load_dotenv
import json

# --- CARGA DE CONFIGURACIÓN ---
load_dotenv()

def cargar_empresas():
    # Intenta cargar el JSON real, si no existe usa el de ejemplo
    archivo = 'empresas.json' if os.path.exists('empresas.json') else 'empresas_example.json'
    with open(archivo, 'r', encoding='utf-8') as f:
        return json.load(f)

USUARIO = os.getenv("CH_USER")
CONTRASEÑA = os.getenv("CH_PASS")
URL_LOGIN = os.getenv("URL_SISTEMA")
URL_REPORTES = os.getenv("CH_URL")

MESES = ["Enero", "Febrero", "Marzo", "Abril"]
SUPERVISORES = cargar_empresas()

def escribir_log(mensaje):
    with open("log_ejecucion.txt", "a", encoding="utf-8") as f:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        f.write(f"[{timestamp}] {mensaje}\n")

def generar_excel_formateado(df, ruta, nombre_empresa):
    try:
        writer = pd.ExcelWriter(ruta, engine='xlsxwriter')
        workbook = writer.book
        sheet_name = 'Reporte'
        
        # --- ESTILOS Y COLORES ---
        paleta = ['#002060', '#FFFF00', '#0070C0', '#FFD966', '#8EA9DB']
        fmt_borde = workbook.add_format({'border': 1, 'border_color': 'black'})
        fmt_total = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#F2F2F2'})
        fmt_header = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9D9D9', 'align': 'center'})

        # 1. TABLA DE REFLEXIONES (CON GRÁFICO)
        df_refl = df[df['Metrica'] == 'Asistencia']
        tabla_refl = df_refl.pivot_table(index='Sucursal', columns='Mes', values='Valor', aggfunc='first').fillna(0)
        orden_meses = [m for m in MESES if m in tabla_refl.columns]
        tabla_refl = tabla_refl.reindex(columns=orden_meses)
        
        tabla_refl.loc['Asistencias Totales'] = tabla_refl.sum()

        worksheet = workbook.add_worksheet(sheet_name)
        worksheet.write(1, 0, 'Sucursal', fmt_header)
        for col_num, mes in enumerate(orden_meses):
            worksheet.write(1, col_num + 1, mes, fmt_header)

        for row_num, (idx, row) in enumerate(tabla_refl.iterrows()):
            fmt = fmt_total if idx == 'Asistencias Totales' else fmt_borde
            worksheet.write(row_num + 2, 0, idx, fmt)
            for col_num, val in enumerate(row):
                worksheet.write(row_num + 2, col_num + 1, val, fmt)

        # --- GRÁFICO ---
        chart = workbook.add_chart({'type': 'column'})
        num_sucursales = len(tabla_refl) - 1 

        for i in range(num_sucursales):
            chart.add_series({
                'name':       [sheet_name, i + 2, 0],
                'categories': [sheet_name, 1, 1, 1, len(orden_meses)],
                'values':     [sheet_name, i + 2, 1, i + 2, len(orden_meses)],
                'fill':       {'color': paleta[i % len(paleta)]},
                'border':     {'color': 'black'},
            })

        chart.set_chartarea({'border': {'none': True}, 'fill': {'none': True}})
        chart.set_plotarea({'border': {'none': True}, 'fill': {'none': True}})
        chart.set_y_axis({'visible': False})
        chart.set_title({'name': f'Reflexiones - {nombre_empresa}', 'name_font': {'bold': True}})
        chart.set_table({'show_keys': True})
        chart.set_legend({'none': True})
        
        worksheet.insert_chart('G2', chart, {'x_scale': 1.3, 'y_scale': 1.1})

        # 2. TABLAS DE CONSEJERÍAS Y VISITAS
        fila_inicio = len(tabla_refl) + 5
        for metrica in ["Consejerias", "Visitas"]:
            df_m = df[df['Metrica'] == metrica]
            if not df_m.empty:
                worksheet.write(fila_inicio, 0, f"Resumen de {metrica}", workbook.add_format({'bold': True}))
                tabla_m = df_m.pivot_table(index='Sucursal', columns='Mes', values='Valor', aggfunc='first').fillna(0)
                tabla_m = tabla_m.reindex(columns=orden_meses)
                
                worksheet.write(fila_inicio + 1, 0, 'Sucursal', fmt_header)
                for col_num, mes in enumerate(orden_meses):
                    worksheet.write(fila_inicio + 1, col_num + 1, mes, fmt_header)
                
                for row_num, (idx, row) in enumerate(tabla_m.iterrows()):
                    worksheet.write(fila_inicio + row_num + 2, 0, idx, fmt_borde)
                    for col_num, val in enumerate(row):
                        worksheet.write(fila_inicio + row_num + 2, col_num + 1, val, fmt_borde)
                
                fila_inicio += len(tabla_m) + 4

        writer.close()
    except Exception as e:
        escribir_log(f"Error generando Excel para {nombre_empresa}: {str(e)}")

def procesar_informes():
    options = webdriver.ChromeOptions()
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 15)

    try:
        driver.get(URL_LOGIN)
        wait.until(EC.presence_of_element_located((By.ID, "user_login"))).send_keys(USUARIO)
        driver.find_element(By.ID, "user_pass").send_keys(CONTRASEÑA)
        driver.find_element(By.ID, "wp-submit").click()
        
        for supervisor, empresas in SUPERVISORES.items():
            if not os.path.exists(supervisor):
                os.makedirs(supervisor)

            for item in empresas:
                # Lógica de Simple vs Grupo
                nombre_reporte = item["nombre_web"] if item["tipo"] == "simple" else item["nombre_reporte"]
                entidades_a_buscar = [item["nombre_web"]] if item["tipo"] == "simple" else item["hijas"]
                
                ruta_final = os.path.join(supervisor, f"{nombre_reporte.replace('/', '_')}.xlsx")

                if os.path.exists(ruta_final):
                    print(f"[-] Saltando {nombre_reporte}: Ya procesado.")
                    continue

                datos_acumulados = []
                sucursales_conocidas = set()
                print(f"\n>>> Procesando Reporte: {nombre_reporte}")

                try:
                    for entidad_web in entidades_a_buscar:
                        for mes in MESES:
                            driver.get(URL_REPORTES)
                            wait.until(EC.presence_of_element_located((By.ID, "reporte_gerencial_empresa")))
                            
                            try:
                                Select(driver.find_element(By.ID, "reporte_gerencial_empresa")).select_by_visible_text(entidad_web)
                            except:
                                escribir_log(f"AVISO: No se encontró '{entidad_web}'.")
                                break

                            Select(driver.find_element(By.ID, "reporte_gerencial_mes")).select_by_visible_text(mes)
                            Select(driver.find_element(By.ID, "reporte_gerencial_anho")).select_by_visible_text("2026")
                            Select(driver.find_element(By.ID, "reporte_gerencial_tipo_vista")).select_by_visible_text("Vista De Impresión")
                            
                            driver.find_element(By.CSS_SELECTOR, "button.btn-primary").click()
                            time.sleep(4) 

                            filas = driver.find_elements(By.CSS_SELECTOR, "table tr")
                            
                            encontrado_en_mes = False
                            for fila in filas[1:]:
                                cols = fila.find_elements(By.TAG_NAME, "td")
                                if len(cols) > 5 and cols[0].text.strip() != "":
                                    # Si es grupo, usamos el nombre de la empresa hija como "Sucursal"
                                    suc = entidad_web if item["tipo"] == "grupo" else cols[0].text.strip()
                                    sucursales_conocidas.add(suc)
                                    
                                    datos_acumulados.append({"Sucursal": suc, "Mes": mes, "Metrica": "Asistencia", "Valor": int(cols[1].text) if cols[1].text.isdigit() else 0})
                                    datos_acumulados.append({"Sucursal": suc, "Mes": mes, "Metrica": "Consejerias", "Valor": int(cols[3].text) if cols[3].text.isdigit() else 0})
                                    datos_acumulados.append({"Sucursal": suc, "Mes": mes, "Metrica": "Visitas", "Valor": int(cols[4].text) if cols[4].text.isdigit() else 0})
                                    encontrado_en_mes = True
                            
                            if not encontrado_en_mes:
                                # Forzado de ceros
                                target_sucs = [entidad_web] if item["tipo"] == "grupo" else (list(sucursales_conocidas) if sucursales_conocidas else ["General"])
                                for s in target_sucs:
                                    for m in ["Asistencia", "Consejerias", "Visitas"]:
                                        datos_acumulados.append({"Sucursal": s, "Mes": mes, "Metrica": m, "Valor": 0})

                    if datos_acumulados:
                        df = pd.DataFrame(datos_acumulados)
                        generar_excel_formateado(df, ruta_final, nombre_reporte)
                        print(f"DONE: {nombre_reporte} guardado.")

                except Exception as e:
                    escribir_log(f"FALLO en '{nombre_reporte}': {str(e)}")

    finally:
        driver.quit()
        print("\n--- PROCESO FINALIZADO ---")

if __name__ == "__main__":
    procesar_informes()