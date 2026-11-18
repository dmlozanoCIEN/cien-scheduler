from playwright.sync_api import sync_playwright
from datetime import datetime
from zoneinfo import ZoneInfo
import os
import zipfile
import glob
import sys
import boto3
from botocore.config import Config

# ============================================================
# ========== CONFIGURACIÓN (viene de GitHub Secrets) =========
# ============================================================

PORTAL_USER = os.environ["PORTAL_USER"]
PORTAL_PASS = os.environ["PORTAL_PASS"]

R2_ACCOUNT_ID = os.environ["R2_ACCOUNT_ID"]
R2_ACCESS_KEY_ID = os.environ["R2_ACCESS_KEY_ID"]
R2_SECRET_ACCESS_KEY = os.environ["R2_SECRET_ACCESS_KEY"]
R2_BUCKET = "cien-informes"

DEPURADOR_URL = "http://app.cien.app/depurador-sesiones-cien.html"

CARPETA_TRABAJO = os.path.join(os.getcwd(), "descargas_cien")
RUTA_CONSOLIDADO = os.path.join(CARPETA_TRABAJO, "CIEN2026.xlsx")
CARPETA_ESTACIONES = os.path.join(CARPETA_TRABAJO, "estaciones")

DOWNLOAD_SECTION = "#app > div > div > div > main > div > div.layout-content-wrapper > main > div > section:nth-child(1)"

CONTENT_TYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Zona horaria de Bogotá: se usa tanto para el navegador como para nombrar archivos
TZ_BOGOTA = ZoneInfo("America/Bogota")


def get_r2_client():
    endpoint_url = f"https://{R2_ACCOUNT_ID}.r2.cloudflarestorage.com"
    return boto3.client(
        "s3",
        endpoint_url=endpoint_url,
        aws_access_key_id=R2_ACCESS_KEY_ID,
        aws_secret_access_key=R2_SECRET_ACCESS_KEY,
        config=Config(signature_version="s3v4"),
        region_name="auto",
    )


def descargar_consolidado_desde_r2(s3):
    print("Descargando consolidado actual desde R2...")
    try:
        s3.download_file(R2_BUCKET, "CIEN2026.xlsx", RUTA_CONSOLIDADO)
        print(f"Consolidado descargado desde R2: {RUTA_CONSOLIDADO}")
    except Exception as e:
        raise Exception(
            f"No se pudo descargar CIEN2026.xlsx desde R2 (bucket '{R2_BUCKET}'). "
            f"Debe existir un consolidado inicial subido manualmente antes del primer "
            f"run automatico. Detalle: {e}"
        )


def hacer_login_con_reintento(page, intentos=3):
    for intento in range(1, intentos + 1):
        try:
            print(f"Intento de login #{intento}...")
            page.goto("https://cien-portal.web.app/login")
            page.fill("#input-0", PORTAL_USER)
            page.fill("#input-2", PORTAL_PASS)
            page.get_by_role("button", name="LOGIN").click()
            page.wait_for_url("**/dashboards/**", timeout=45000)
            print("Login exitoso.")
            return
        except Exception as e:
            print(f"Fallo el intento {intento}: {e}")
            if intento == intentos:
                raise Exception("No se pudo iniciar sesion despues de varios intentos.")
            page.wait_for_timeout(5000)


def exportar_sesiones_del_dia(page):
    print("Abriendo portal...")
    hacer_login_con_reintento(page)

    print("Navegando a la lista de sesiones...")
    page.goto("https://cien-portal.web.app/sessions/list?updated_at=today")
    page.wait_for_load_state("networkidle")

    print("Haciendo clic en el boton de descarga...")
    botones = page.locator(f"{DOWNLOAD_SECTION} button")
    botones.nth(2).click()
    page.wait_for_timeout(2000)

    print("Haciendo clic en 'Export as Excel' y capturando la descarga...")
    with page.expect_download(timeout=30000) as download_info:
        page.get_by_text("Export as Excel", exact=False).click(timeout=20000)

    download = download_info.value
    # Fecha calculada en hora de Bogotá, no en hora del runner (UTC)
    fecha = datetime.now(TZ_BOGOTA).strftime("%Y-%m-%d")
    ruta_dia = os.path.join(CARPETA_TRABAJO, f"sesiones_{fecha}.xlsx")
    download.save_as(ruta_dia)
    print(f"Export del dia guardado: {ruta_dia}")
    return ruta_dia


def fusionar_consolidado(page, ruta_dia):
    print(f"Abriendo herramienta: {DEPURADOR_URL}")
    page.goto(DEPURADOR_URL)
    page.wait_for_load_state("networkidle")

    print("Seleccionando modo 'Fusionar consolidado + dia'...")
    page.check("#modeMerge", force=True)

    print("Cargando archivo A (consolidado existente)...")
    page.set_input_files("#fileA", RUTA_CONSOLIDADO)

    print("Cargando archivo B (sesiones del dia)...")
    page.set_input_files("#fileB", ruta_dia)

    print("Esperando que el boton 'Fusionar' se habilite...")
    page.wait_for_selector("#btnMerge:not([disabled])", timeout=30000)

    error_box = page.locator("#mErr")
    if error_box.is_visible():
        raise Exception(f"Error reportado en Fusionar: {error_box.inner_text()}")

    print("Haciendo clic en 'Fusionar y renumerar'...")
    page.click("#btnMerge")

    print("Esperando que el boton de descarga se habilite...")
    page.wait_for_selector("#btnMergeXlsx:not([disabled])", timeout=30000)

    print("Descargando consolidado actualizado...")
    with page.expect_download(timeout=30000) as download_info:
        page.click("#btnMergeXlsx")

    download = download_info.value
    download.save_as(RUTA_CONSOLIDADO)
    print(f"Consolidado actualizado: {RUTA_CONSOLIDADO}")


def dividir_por_estacion(page):
    print("Seleccionando modo 'Dividir por estacion'...")
    page.check("#modeSplit", force=True)

    print("Cargando el consolidado actualizado para dividir...")
    page.set_input_files("#file", RUTA_CONSOLIDADO)

    print("Esperando que la division se procese...")
    try:
        page.wait_for_selector("#splitExports:not([hidden])", timeout=40000)
    except Exception:
        error_box = page.locator("#splitErr")
        if error_box.is_visible():
            raise Exception(f"Error reportado en Dividir por estacion: {error_box.inner_text()}")
        raise

    error_box = page.locator("#splitErr")
    if error_box.is_visible():
        raise Exception(f"Error reportado en Dividir por estacion: {error_box.inner_text()}")

    print("Descargando ZIP por estaciones...")
    with page.expect_download(timeout=30000) as download_info:
        page.click("#btnSplitZip")

    download = download_info.value
    ruta_zip = os.path.join(CARPETA_TRABAJO, "CIEN2026-estaciones.zip")
    download.save_as(ruta_zip)
    print(f"ZIP de estaciones descargado: {ruta_zip}")
    return ruta_zip


def descomprimir_estaciones(ruta_zip):
    print(f"Descomprimiendo en: {CARPETA_ESTACIONES}")
    os.makedirs(CARPETA_ESTACIONES, exist_ok=True)

    for f in glob.glob(f"{CARPETA_ESTACIONES}/*.xlsx"):
        os.remove(f)

    with zipfile.ZipFile(ruta_zip, "r") as z:
        z.extractall(CARPETA_ESTACIONES)

    print(f"Archivos por estacion descomprimidos en: {CARPETA_ESTACIONES}")


def subir_a_r2(s3):
    print(f"Subiendo consolidado a la raiz del bucket '{R2_BUCKET}'...")
    s3.upload_file(
        RUTA_CONSOLIDADO,
        R2_BUCKET,
        "CIEN2026.xlsx",
        ExtraArgs={"ContentType": CONTENT_TYPE_XLSX},
    )
    print("CIEN2026.xlsx subido/actualizado en la raiz.")

    archivos_estacion = glob.glob(f"{CARPETA_ESTACIONES}/*.xlsx")
    print(f"Subiendo {len(archivos_estacion)} archivos de estaciones a la raiz del bucket...")
    for ruta_archivo in archivos_estacion:
        nombre = os.path.basename(ruta_archivo)
        s3.upload_file(
            ruta_archivo,
            R2_BUCKET,
            nombre,
            ExtraArgs={"ContentType": CONTENT_TYPE_XLSX},
        )
        print(f"  -> {nombre} subido/actualizado")

    print("Todos los archivos subidos/actualizados en la raiz del bucket.")


def proceso_completo():
    os.makedirs(CARPETA_TRABAJO, exist_ok=True)
    s3 = get_r2_client()

    descargar_consolidado_desde_r2(s3)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            accept_downloads=True,
            viewport={"width": 1366, "height": 900},
            timezone_id="America/Bogota",
            locale="es-CO",
        )
        page = context.new_page()

        try:
            ruta_dia = exportar_sesiones_del_dia(page)
            fusionar_consolidado(page, ruta_dia)
            ruta_zip = dividir_por_estacion(page)
        except Exception:
            page.screenshot(path=os.path.join(CARPETA_TRABAJO, "error_screenshot.png"))
            raise
        finally:
            browser.close()

    descomprimir_estaciones(ruta_zip)
    subir_a_r2(s3)
    print("Proceso completo finalizado con exito: archivos subidos a R2.")


if __name__ == "__main__":
    try:
        proceso_completo()
    except Exception as e:
        print(f"ERROR: {e}")
        sys.exit(1)
