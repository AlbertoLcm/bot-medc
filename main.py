import asyncio
import os
import shutil
import glob
import time
import pandas as pd
from playwright.async_api import async_playwright

# --- 1. CONFIGURACIÓN PARA DESCARGAS AUTOMÁTICAS ---
download_dir = os.path.join(os.getcwd(), "descargas_pdf")

if os.path.exists(download_dir):
    shutil.rmtree(download_dir)
os.makedirs(download_dir)
# --- FIN DE LA CONFIGURACIÓN ---

oficios_df = pd.read_excel("layout.xlsx", dtype=str)
oficios_df["Periodo Requerido"] = pd.to_datetime(oficios_df["Periodo Requerido"])
oficios_df["Periodo mes"] = (oficios_df["Periodo Requerido"].dt.to_period("M").dt.to_timestamp())
oficios_dict = oficios_df.to_dict(orient="records")

cuentas = {}
for oficio in oficios_dict:
    cuenta = oficio["Numero de cuenta"]
    periodo = oficio["Periodo mes"]
    if cuenta not in cuentas:
        cuentas[cuenta] = []
    cuentas[cuenta].append(periodo)

if os.path.exists("Resultados.xlsx"):
    os.remove("Resultados.xlsx")

async def consulta_cuenta_y_descarga(page, context, resultados, page_current, cuenta, periodos_consultados, download_dir):
    try:
        await page.goto("https://150.100.42.196:36060/servAutMEDC/content/pages/Aclaraciones/ACL.xhtml", wait_until="networkidle")

        input_selector = "#formAutor\\:j_id387930406_3a861378"
        await page.wait_for_selector(input_selector)
        input_cuenta = page.locator(input_selector)
        await input_cuenta.fill("") # clear
        await input_cuenta.fill(cuenta)
        await input_cuenta.press("Enter")

        await page.wait_for_selector("#formAutor\\:tableACLA_paginator_top")
        select_selector = "#formAutor\\:tableACLA\\:j_id27"
        await page.wait_for_selector(select_selector)
        await page.select_option(select_selector, label="100")
        await asyncio.sleep(2) # Allow time for table reload after select
        
        await page.wait_for_selector(".ui-paginator-pages")
        
        paginator_pages = page.locator(".ui-paginator-pages a")
        num_pages = await paginator_pages.count()

        for pagina in range(num_pages):
            try:
                if page_current < pagina:
                    continue

                if pagina > 0:
                    # Re-locate paginator elements in case DOM changed
                    paginator_pages = page.locator(".ui-paginator-pages a")
                    await paginator_pages.nth(pagina).click()
                    await asyncio.sleep(2) # Wait for page load

                await page.wait_for_selector("#formAutor\\:tableACLA_data")
                
                # Buscamos las filas con el atributo data-ri
                filas = page.locator("#formAutor\\:tableACLA_data tr[data-ri]")
                num_filas = await filas.count()

                page_current = pagina
                contador_archivos = 0

                if num_filas > 0:
                    for idx in range(num_filas):
                        fila = filas.nth(idx)
                        contador_archivos += 1
                        
                        columnas = fila.locator("td")
                        cuenta_texto = await columnas.nth(0).inner_text()
                        cuenta_texto = cuenta_texto.strip()
                        cliente = await columnas.nth(1).inner_text()
                        periodo = await columnas.nth(2).inner_text()
                        corte = await columnas.nth(3).inner_text()
                        corte = corte.strip()
                        
                        enlace_PDF = columnas.nth(4).locator("td button").first

                        ID_periodo = cuenta_texto + corte

                        if ID_periodo in periodos_consultados:
                            continue

                        corte_fecha = pd.to_datetime(corte)
                        periodo_fecha = corte_fecha.to_period("M").to_timestamp()

                        if periodo_fecha not in cuentas[cuenta_texto]:
                            continue

                        archivos_antes_del_clic = glob.glob(os.path.join(download_dir, "*.pdf"))

                        # In Playwright, we wait for the 'download' event while clicking
                        try:
                             async with page.expect_download(timeout=20_000) as download_info:
                                await enlace_PDF.click()
                             download = await download_info.value
                             
                             # Specify custom path or use generated name
                             download_path = os.path.join(download_dir, download.suggested_filename)
                             await download.save_as(download_path)
                             
                             archivo_descargado = True
                             nombre_archivo_descargado = download_path
                        except Exception as e:
                             archivo_descargado = False
                             print(f"Exception during download: {e}")

                        periodos_consultados.append(ID_periodo)

                        if archivo_descargado:
                            resultados.append([cuenta_texto, corte_fecha, "Correcto"])
                            print(f"✅ Periodo descargado CORRECTAMENTE: {corte_fecha}. Archivo: {os.path.basename(nombre_archivo_descargado)}")
                        else:
                            resultados.append([cuenta_texto, corte_fecha, "Error de MEDC"])
                            print(f"❌ Error al descargar {corte_fecha}.")
                            # Recursive call on error, similar to original script
                            await consulta_cuenta_y_descarga(page, context, resultados, page_current, cuenta, periodos_consultados, download_dir)

            except Exception as e:
                print(f"Error within page iteration loop: {e}")
                await consulta_cuenta_y_descarga(page, context, resultados, page_current, cuenta, periodos_consultados, download_dir)
                break

    except Exception as e:
        print(f"Error in main interaction flow: {e}")
        # Note: Depending on the error, infinite recursion is possible here as in the original script. 
        # Consider adding a max_retries limit if needed.

async def main():
    async with async_playwright() as p:
        # Launch Chromium. Equivalent options for ignoring SSL errors and downloads
        browser = await p.chromium.launch(
            channel="chrome",
            headless=False, # Set to True if you don't want to see the UI
            args=['--ignore-certificate-errors', '--ignore-ssl-errors=yes']
        )
        
        # In Playwright, downloads are handled at the context level or via events
        context = await browser.new_context(
            ignore_https_errors=True,
            accept_downloads=True
        )
        
        page = await context.new_page()

        contador = 0
        resultados = []
        page_current = 0
        periodos_consultados = []

        for cuenta in cuentas.keys():
            contador += 1
            print(f"\n--- Procesando cuenta: {cuenta} de {len(cuentas[cuenta])} periodos requeridos ---")
            await consulta_cuenta_y_descarga(page, context, resultados, page_current, cuenta, periodos_consultados, download_dir)

        await browser.close()

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

if __name__ == "__main__":
    asyncio.run(main())