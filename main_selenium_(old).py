from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import re
import os
import time
import glob
import shutil

# --- 1. CONFIGURACIÓN PARA DESCARGAS AUTOMÁTICAS ---
# Define una carpeta para las descargas dentro del directorio actual del script
if os.path.exists("descargas_pdf"):
    shutil.rmtree("descargas_pdf")

download_dir = os.path.join(os.getcwd(), "descargas_pdf")
os.makedirs(download_dir)

# Configura Chrome para que use esa carpeta y descargue los PDF directamente
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,  # Para no preguntar dónde guardar
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True,  # MUY IMPORTANTE: Evita que Chrome abra el PDF
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--ignore-ssl-errors=yes")
options.add_argument("--ignore-certificate-errors")
# --- FIN DE LA CONFIGURACIÓN ---

oficios_df = pd.read_excel("LAY OUT.xlsx", dtype=str)
oficios_df["Periodo Requerido"] = pd.to_datetime(oficios_df["Periodo Requerido"])
oficios_df["Periodo mes"] = (oficios_df["Periodo Requerido"].dt.to_period("M").dt.to_timestamp())
oficios_dict = oficios_df.to_dict(orient="records")

# Agrupamos todos los numeros de cuenta, y en un array almacenamos todos las fechas
cuentas = {}
for oficio in oficios_dict:
    cuenta = oficio["Numero de cuenta"]
    periodo = oficio["Periodo mes"]
    if cuenta not in cuentas:
        cuentas[cuenta] = []
    cuentas[cuenta].append(periodo)

if os.path.exists("Resultados.xlsx"):
    os.remove("Resultados.xlsx")


def consulta_cuenta_y_descarga(driver, wait, resultados, page_current, cuenta, periodos_consultados):
    driver.get("https://150.100.42.196:36060/servAutMEDC/content/pages/Aclaraciones/ACL.xhtml")

    wait.until(EC.presence_of_element_located((By.ID, "formAutor:j_id387930406_3a861378")))
    # Ingresamos la cuenta
    input_cuenta = driver.find_element(By.ID, "formAutor:j_id387930406_3a861378")
    input_cuenta.clear()
    input_cuenta.send_keys(cuenta)
    input_cuenta.send_keys(Keys.ENTER)
    # Esperamos a que se muestren los resultados
    # Obtenemos la paginación
    wait.until(EC.presence_of_element_located((By.ID, "formAutor:tableACLA_paginator_top")))
    wait.until(EC.presence_of_element_located((By.ID, "formAutor:tableACLA:j_id27")))
    select_element = Select(driver.find_element(By.ID, "formAutor:tableACLA:j_id27"))
    select_element.select_by_visible_text("100")
    time.sleep(2)
    # Esperamos a que se cargue el contenedor de la tabla de resultados
    wait.until(EC.presence_of_element_located((By.CLASS_NAME, "ui-paginator-pages")))

    container_paginator = driver.find_element(By.CLASS_NAME, "ui-paginator-pages")
    paginas = container_paginator.find_elements(By.TAG_NAME, "a")

    # Realizamos la iteración de las páginas

    for pagina in range(len(paginas)):
        try:
            # Si la pagina no es la page_current omitimos la pagina
            if page_current < pagina:
                continue

            if pagina > 0:
                paginas = container_paginator.find_elements(By.TAG_NAME, "a")
                paginas[pagina].click()
                time.sleep(2)  # Espera para que se cargue la página

            wait.until(
                EC.presence_of_element_located((By.ID, "formAutor:tableACLA_data"))
            )
            tbody = driver.find_element(By.ID, "formAutor:tableACLA_data")
            # Buscamos las filas con el atributo data-ri, ya que hay mas tablas dentro de tbody
            filas = tbody.find_elements(By.XPATH, ".//tr[@data-ri]")

            # Asignamos la pagina actual como page_current
            page_current = pagina

            contador_archivos = 0

            if tbody and len(filas) > 1:
                for fila in filas:
                    contador_archivos += 1
                    columnas = fila.find_elements(By.TAG_NAME, "td")
                    cuenta = columnas[0].text.strip()
                    cliente = columnas[1].text.strip()
                    periodo = columnas[2].text.strip()
                    corte = columnas[3].text.strip()  # Formato: 2025-09-30 00:00:00.0
                    enlaces = columnas[4].find_elements(By.TAG_NAME, "td")
                    enlace_PDF = enlaces[0].find_element(By.TAG_NAME, "button")

                    ID_periodo = cuenta+corte

                    # Validamos si esta cuenta y periodo no han sido consultadas
                    if ID_periodo in periodos_consultados:
                        # Si ya fue consultado lo omitimos
                        continue

                    # Si el periodo (columna corte) de la fila está en la lista de periodos requeridos para esta cuenta, descargamos el PDF
                    # periodo_fecha = pd.to_datetime(corte, format='%Y-%m-%d %H:%M:%S.%f')
                    corte_fecha = pd.to_datetime(corte)
                    periodo_fecha = corte_fecha.to_period("M").to_timestamp()

                    if periodo_fecha not in cuentas[cuenta]:
                        continue

                    # Antes del clic, lista los archivos PDF existentes en el directorio de descargas
                    archivos_antes_del_clic = glob.glob(
                        os.path.join(download_dir, "*.pdf")
                    )

                    enlace_PDF.click()

                    # Tiempo de espera en la descarga
                    tiempo_espera_maximo = 10  # segundos
                    tiempo_inicio = time.time()
                    archivo_descargado = False
                    nombre_archivo_descargado = None

                    while (time.time() - tiempo_inicio) < tiempo_espera_maximo:
                        # 3. Lista los archivos después del clic
                        archivos_despues_del_clic = glob.glob(
                            os.path.join(download_dir, "*.pdf")
                        )

                        # 4. Compara las listas para encontrar un archivo nuevo
                        archivos_nuevos = [
                            f
                            for f in archivos_despues_del_clic
                            if f not in archivos_antes_del_clic
                        ]

                        if archivos_nuevos:
                            archivo_descargado = True
                            nombre_archivo_descargado = archivos_nuevos[0]  # Tomamos el primer archivo nuevo
                            break  # ¡Se encontró el archivo, salimos del bucle!

                        time.sleep(0.5)

                    # --- Resultado de la verificación ---

                    # Agregamos esta cuenta y periodo como "consultada"
                    periodos_consultados.append(ID_periodo)

                    if archivo_descargado:
                        resultados.append([cuenta, corte_fecha, "Correcto"])
                        print(f"✅ Periodo descargado CORRECTAMENTE: {corte_fecha}. Archivo: {os.path.basename(nombre_archivo_descargado)}")

                    else:
                        resultados.append([cuenta, corte_fecha, "Error de MEDC"])
                        print(f"❌ Error al descargar {corte_fecha}. No se encontró un nuevo PDF en {download_dir} después de {tiempo_espera_maximo}s.")
                        consulta_cuenta_y_descarga(driver, wait, resultados, page_current, cuenta, periodos_consultados)

        except Exception as e:
            consulta_cuenta_y_descarga(driver, wait, resultados, page_current, cuenta, periodos_consultados)
            break



# ---- CÓDIGO PRINCIPAL ----
# Inicializa el driver CON las opciones de descarga configuradas
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

contador = 0
resultados = []
page_current = 0

# Varible donde almacenaremos todos los EDC consultados
periodos_consultados = [] # Se guarda de acuerdo a un ID generado de la tabla de periodos

for cuenta in cuentas.keys():
    contador += 1
    print(f"\n--- Procesando cuenta: {cuenta} de {len(cuentas[cuenta])} periodos requeridos ---")
    consulta_cuenta_y_descarga(driver, wait, resultados, page_current, cuenta, periodos_consultados)

driver.quit()
print("\n--- Todas las cuentas han sido consultadas ---")
oficios_df['Periodo mes'] = (oficios_df["Periodo Requerido"].dt.to_period("M").dt.to_timestamp()).dt.strftime("%Y-%m-%d")
oficios_df['ID'] = oficios_df['Numero de cuenta']+'-'+oficios_df['Periodo mes']

resultados_df = pd.DataFrame(resultados, columns=["Cuenta", "Periodo Descargado", "Estatus"])
resultados_df['Periodo mes'] = (resultados_df["Periodo Descargado"].dt.to_period("M").dt.to_timestamp()).dt.strftime("%Y-%m-%d")
resultados_df['ID'] = resultados_df['Cuenta']+'-'+resultados_df['Periodo mes']

resultados_total = pd.merge(oficios_df, resultados_df[['ID', 'Periodo Descargado', 'Estatus']], on='ID', how='left')
resultados_total = resultados_total[['Numero de cuenta', 'Periodo Requerido', 'Periodo Descargado', 'Estatus']]
resultados_total = resultados_total.replace(pd.NaT, 'No Encontrado')

print("--- Los resultados han sido exportados en un excel ---")
resultados_total.to_excel("Resultados Consultados.xlsx", index=False)