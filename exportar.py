import os
import re
import sys
import zipfile
from datetime import datetime
from zoneinfo import ZoneInfo

import boto3
from botocore.client import Config
from playwright.sync_api import sync_playwright

# ============================================================
# CONFIGURACION
# ============================================================

PORTAL_LOGIN_URL = "https://cien-portal.web.app/login"
PORTAL_SESSIONS_URL = "https://cien-portal.web.app/sessions/list?updated_at=today"
DEPURADOR_URL = "http://app.cien.app/depurador-sesiones-cien.html"

PORTAL_USER = os.environ.get("PORTAL_USER")
PORTAL_PASS = os.environ.get("PORTAL_PASS")

R2_ACCOUNT_ID = os.environ.get("R2_ACCOUNT_ID")
R2_ACCESS_KEY_ID = os.environ.get("R2_ACCESS_KEY_ID")
R2_SECRET_ACCESS_KEY = os.environ.get("R2_SECRET_ACCESS_KEY")
R2_BUCKET = os.environ.get("R2_BUCKET", "cien-informes")

CARPETA_TRABAJO = "descargas_cien"
CARPETA_ESTACIONES = os.path.join(CARPETA_TRABAJO, "estaciones")
RUTA_CONSOLIDADO = os.path.join(CARPETA_TRABAJO, "CIEN2026.xlsx")

TZ_BOGOTA = ZoneInfo("America/Bogota")

MODO_HEADLESS = True
TIMEOUT_LOGIN_MS = 45000
TIMEOUT_EXPORT_MS = 30000
TIMEOUT_PROCESO_GRANDE_MS = 120000
TIMEOUT_DIVISION_TOTAL_MS = 300000
INTERVALO_POLL_MS = 15000

DOWNLOAD_SECTION = (
    "#app > div > div > div > main > div > div.layout-content-wrapper "
    "> main > div > section:nth-child(1)"
)


def validar_configuracion():
    faltantes = []
    if not PORTAL_USER:
        faltantes.append("PORTAL_USER")
    if not PORTAL_PASS:
        faltantes.append("PORTAL_PASS")
    if not R2_ACCOUNT_ID:
        faltantes.append("R2_ACCOUNT_ID")
    if not R2_ACCESS_KEY_ID:
        faltantes.append("R2_ACCESS_KEY_ID")
    if not R2_SECRET_ACCESS_KEY:
        faltantes.append("R2_SECRET_ACCESS_KEY")
    if faltantes:
        print(f"ERROR: faltan variables de entorno/secrets: {', '.join(faltantes)}")
        sys.exit(1)


def cliente_r2():
    endpoint = f"https://{R2_ACCOUNT_ID}.r2.cloudflarestorage.com"
    return boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=R2_ACCESS_KEY_ID,
        aws_secret_access_key=R2_SECRET_ACCESS_KEY,
        config=Config(signature_version="s3v4"),
        region_name="auto",
    )


def descargar_consolidado_desde_r2():
    print("Descargando consolidado actual desde R2...")
    os.makedirs(CARPETA_TRABAJO, exist_ok=True)
    s3 = cliente_r2()
    s3.download_file(R2_BUCKET, "CIEN2026.xlsx", RUTA_CONSOLIDADO)
    print(f"Consolidado descargado desde R2: {RUTA_CONSOLIDADO}")


def subir_a_r2():
    print("Subiendo archivos a la raiz del bucket R2...")
    s3 = cliente_r2()
    archivos_a_subir = []

    if os.path.exists(RUTA_CONSOLIDADO):
        archivos_a_subir.append(RUTA_CONSOLIDADO)

    if os.path.isdir(CARPETA_ESTACIONES):
        for nombre in os.listdir(CARPETA_ESTACIONES):
            if nombre.lower().endswith(".xlsx"):
                archivos_a_subir.append(os.path.join(CARPETA_ESTACIONES, nombre))

    for ruta in archivos_a_subir:
        key = os.path.basename(ruta)
        s3.upload_file(
            ruta,
            R2_BUCKET,
            key,
            ExtraArgs={
                "ContentType": (
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                )
            },
        )
        print(f"  Subido: {key}")

    print("Subida a R2 completada.")


def conectar_diagnostico_js(page):
    page.on(
        "pageerror",
        lambda exc: print(f"[JS ERROR] {exc}"),
    )
    page.on(
        "console",
        lambda msg: print(f"[JS CONSOLE:{msg.type}] {msg.text}")
        if msg.type == "error"
        else None,
    )


def asignar_archivo_y_disparar_change(page, selector, ruta_archivo):
    """
    Asigna un archivo a un input[type=file] y fuerza manualmente el
    evento 'change', ya que Playwright puede no disparar el evento
    si la ruta asignada es identica a una asignacion previa en el
    mismo input, aunque el contenido del archivo en disco haya cambiado.
    """
    page.set_input_files(selector, ruta_archivo)
    page.eval_on_selector(
        selector,
        "el => el.dispatchEvent(new Event('change', { bubbles: true }))"
    )


def hacer_login_con_reintento(page, intentos=3):
    for intento in range(1, intentos + 1):
        try:
            print(f"Intento de login #{intento}...")
            page.goto(PORTAL_LOGIN_URL, timeout=TIMEOUT_LOGIN_MS)
            page.wait_for_selector("#input-0", timeout=TIMEOUT_LOGIN_MS)
            page.fill("#input-0", PORTAL_USER)
            page.fill("#input-2", PORTAL_PASS)
            page.get_by_role("button", name="LOGIN").click()
            page.wait_for_url("**/dashboards/**", timeout=TIMEOUT_LOGIN_MS)
            print("Login exitoso.")
            return
        except Exception as e:
            print(f"Fallo intento de login #{intento}: {e}")
            if intento == intentos:
                page.screenshot(path="error_screenshot.png")
                raise
            page.wait_for_timeout(5000)


def exportar_sesiones_del_dia(page):
    print("Navegando a la lista de sesiones...")
    page.goto(PORTAL_SESSIONS_URL, timeout=TIMEOUT_LOGIN_MS)
    page.wait_for_load_state("networkidle")

    print("Haciendo clic en el boton de descarga...")
    botones = page.locator(f"{DOWNLOAD_SECTION} button")
    botones.nth(2).click()

    print("Haciendo clic en 'Export as Excel' y capturando la descarga...")
    with page.expect_download(timeout=TIMEOUT_EXPORT_MS) as download_info:
        page.get_by_text("Export as Excel", exact=False).click(timeout=TIMEOUT_EXPORT_MS)
    download = download_info.value

    fecha = datetime.now(TZ_BOGOTA).strftime("%Y-%m-%d")
    ruta_dia = os.path.join(CARPETA_TRABAJO, f"sesiones_{fecha}.xlsx")
    download.save_as(ruta_dia)
    print(f"Export del dia guardado: {ruta_dia}")
    return ruta_dia


def abrir_depurador(page):
    print(f"Abriendo herramienta: {DEPURADOR_URL}")
    page.goto(DEPURADOR_URL, timeout=TIMEOUT_LOGIN_MS)
    page.wait_for_load_state("networkidle")
    page.wait_for_selector("#modeSingle", timeout=30000)


def fusionar_consolidado(page, ruta_dia):
    print("Seleccionando modo 'Fusionar consolidado + dia'...")
    page.check("#modeMerge", force=True)
    page.wait_for_timeout(500)

    print("Cargando archivo A (consolidado existente)...")
    asignar_archivo_y_disparar_change(page, "#fileA", RUTA_CONSOLIDADO)

    print("Cargando archivo B (sesiones del dia)...")
    asignar_archivo_y_disparar_change(page, "#fileB", ruta_dia)

    print("Esperando que el boton 'Fusionar' se habilite...")
    page.wait_for_selector("#btnMerge:not([disabled])", timeout=TIMEOUT_PROCESO_GRANDE_MS)

    print("Haciendo clic en 'Fusionar y renumerar'...")
    page.click("#btnMerge")

    print("Esperando que el boton de descarga se habilite...")
    page.wait_for_selector(
        "#btnMergeXlsx:not([disabled])", timeout=TIMEOUT_PROCESO_GRANDE_MS
    )

    print("Descargando consolidado fusionado (aun sin depurar)...")
    with page.expect_download(timeout=TIMEOUT_PROCESO_GRANDE_MS) as download_info:
        page.click("#btnMergeXlsx")
    download = download_info.value
    download.save_as(RUTA_CONSOLIDADO)
    print(f"Consolidado fusionado guardado: {RUTA_CONSOLIDADO}")


def depurar_consolidado(page):
    print("Seleccionando modo 'Depurar'...")
    page.check("#modeSingle", force=True)
    page.wait_for_timeout(500)

    print("Cargando el consolidado fusionado para depurar...")
    asignar_archivo_y_disparar_change(page, "#file", RUTA_CONSOLIDADO)

    print("Esperando que la depuracion se procese...")
    page.wait_for_selector(
        "#results:not([hidden]), #fileErr:not([hidden])",
        timeout=TIMEOUT_PROCESO_GRANDE_MS,
    )

    if page.locator("#fileErr:not([hidden])").count() > 0:
        mensaje_error = page.locator("#fileErr").inner_text()
        page.screenshot(path="error_screenshot.png")
        raise RuntimeError(f"Error reportado por la herramienta al depurar: {mensaje_error}")

    print("Esperando que la zona de descarga este disponible...")
    page.wait_for_selector(
        "#singleExports:not([hidden])", timeout=TIMEOUT_PROCESO_GRANDE_MS
    )

    print("Descargando consolidado ya depurado (.xlsx)...")
    with page.expect_download(timeout=TIMEOUT_PROCESO_GRANDE_MS) as download_info:
        page.click("#btnXlsx")
    download = download_info.value
    download.save_as(RUTA_CONSOLIDADO)
    print(f"Consolidado depurado guardado: {RUTA_CONSOLIDADO}")


def dividir_por_estacion(page):
    print("Seleccionando modo 'Dividir por estacion'...")
    page.check("#modeSplit", force=True)
    page.wait_for_timeout(500)

    print("Cargando el consolidado ya depurado para dividir...")
    asignar_archivo_y_disparar_change(page, "#file", RUTA_CONSOLIDADO)

    print("Esperando que la division se procese...")
    transcurrido_ms = 0
    resultado = None

    while transcurrido_ms < TIMEOUT_DIVISION_TOTAL_MS:
        if page.locator("#splitExports:not([hidden])").count() > 0:
            resultado = "exito"
            break
        if page.locator("#splitErr:not([hidden])").count() > 0:
            resultado = "error_split"
            break
        if page.locator("#fileErr:not([hidden])").count() > 0:
            resultado = "error_file"
            break
        page.wait_for_timeout(INTERVALO_POLL_MS)
        transcurrido_ms += INTERVALO_POLL_MS
        print(f"  ...sigue procesando ({transcurrido_ms // 1000}s transcurridos)")

    if resultado is None:
        page.screenshot(path="error_screenshot.png")
        raise RuntimeError(
            f"Timeout: la division no termino en {TIMEOUT_DIVISION_TOTAL_MS // 1000}s. "
            "Revisa error_screenshot.png."
        )

    if resultado == "error_split":
        mensaje_error = page.locator("#splitErr").inner_text()
        page.screenshot(path="error_screenshot.png")
        raise RuntimeError(f"Error reportado por la herramienta al dividir: {mensaje_error}")

    if resultado == "error_file":
        mensaje_error = page.locator("#fileErr").inner_text()
        page.screenshot(path="error_screenshot.png")
        raise RuntimeError(f"Error al cargar el archivo para dividir: {mensaje_error}")

    print("Division completada. Descargando ZIP por estacion...")
    with page.expect_download(timeout=TIMEOUT_PROCESO_GRANDE_MS) as download_info:
        page.click("#btnSplitZip")
    download = download_info.value
    ruta_zip = os.path.join(CARPETA_TRABAJO, "CIEN2026-estaciones.zip")
    download.save_as(ruta_zip)
    print(f"ZIP de estaciones guardado: {ruta_zip}")
    return ruta_zip


def descomprimir_zip(ruta_zip):
    print(f"Descomprimiendo {ruta_zip}...")
    os.makedirs(CARPETA_ESTACIONES, exist_ok=True)
    with zipfile.ZipFile(ruta_zip, "r") as z:
        z.extractall(CARPETA_ESTACIONES)
    print(f"Archivos de estaciones extraidos en: {CARPETA_ESTACIONES}")


def proceso_completo():
    descargar_consolidado_desde_r2()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=MODO_HEADLESS)
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 900},
            timezone_id="America/Bogota",
            locale="es-CO",
        )
        page = context.new_page()
        conectar_diagnostico_js(page)

        try:
            print("Abriendo portal...")
            hacer_login_con_reintento(page)

            ruta_dia = exportar_sesiones_del_dia(page)

            abrir_depurador(page)
            fusionar_consolidado(page, ruta_dia)
            depurar_consolidado(page)
            ruta_zip = dividir_por_estacion(page)
            descomprimir_zip(ruta_zip)

        except Exception as e:
            print(f"ERROR: {e}")
            try:
                page.screenshot(path="error_screenshot.png")
            except Exception:
                pass
            browser.close()
            raise

        browser.close()

    subir_a_r2()
    print("Proceso completo finalizado correctamente.")


if __name__ == "__main__":
    validar_configuracion()
    proceso_completo()
