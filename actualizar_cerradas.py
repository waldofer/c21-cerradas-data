"""
Actualizador semanal del dataset de cerradas - Century 21 Paraguay
====================================================================
Corre los lunes (Task Scheduler), exporta lo cerrado en 21online desde la
última actualización hasta hoy, y actualiza (upsert por ID) el dataset
maestro en el repo de GitHub c21-cerradas-data.

REQUIERE ENGANCHARSE A TU LOGIN EXISTENTE:
Reemplazá la función `login()` por la misma que ya usás en infocasas-bot
(scrape.py) para entrar a 21online.century21.com.py. No la reescribas,
reusala tal cual — esto vive en un repo aparte, no toques infocasas-bot.

Flujo:
1. Lee data/cerradas.json del repo local (el dataset acumulado hasta ahora)
2. Calcula la fecha de corte: la fecha de cierre más reciente que ya tenemos,
   menos 14 días de margen (para agarrar cierres que se cargaron tarde)
3. Entra a 21online, exporta el reporte de cerradas filtrado por ESE rango
   (fecha_corte -> hoy)
4. Parsea el Excel descargado (misma lógica de limpieza que ya usamos:
   USD/PYG, LAT/LON, PRECIO vs PRECIO CIERRE, FECHA ALTA)
5. Upsert por ID sobre el dataset acumulado (los IDs nuevos se agregan,
   los existentes se actualizan si cambiaron)
6. Commit + push al repo c21-cerradas-data

Uso:
    python actualizar_cerradas.py

Requiere: playwright, openpyxl, git configurado con acceso de push al repo.
"""

import json
import re
import subprocess
import os
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

REPO_DIR = Path(__file__).parent  # asumido: este script vive adentro del repo c21-cerradas-data
DATASET_PATH = REPO_DIR / "data" / "cerradas.json"
DESCARGA_TMP = REPO_DIR / "_tmp_export.xlsx"

TC_APROX = 6100  # Gs por USD, fallback si algún PRECIO viene en PYG sin más contexto
MARGEN_DIAS = 14  # relee los últimos N días por si hubo cargas tardías


BASE_URL = "https://21online.century21.com.py"


def construir_url_busqueda(fecha_desde: str, fecha_hasta: str) -> str:
    params = {
        "oficinaStatus": "todas",
        "oficina": "",
        "estatus": "cerrada",
        "tipoOperacion": "venta",
        "mesCierreDesde": fecha_desde,
        "mesCierreHasta": fecha_hasta,
        "cantidad": "30000",
        "conFoto": "1",
    }
    query = "&".join(f"{k}={v}" for k, v in params.items())
    return f"{BASE_URL}/inventarios?{query}"


def login(page, max_intentos: int = 3):
    """
    Login en 21online.century21.com.py. Adaptado de la función _login() de
    infocasas-bot/scrape.py (mismo sitio, mismo formulario): llena los dos
    primeros inputs visibles del form (usuario/contraseña) y clickea el
    submit, con reintentos si la sesión sigue en /login.

    Credenciales SIEMPRE desde variables de entorno, nunca hardcodeadas ni
    pegadas en el chat:
        C21_USUARIO, C21_PASSWORD
    """
    usuario = os.environ["C21_USUARIO"]
    password = os.environ["C21_PASSWORD"]

    for intento in range(1, max_intentos + 1):
        page.goto(f"{BASE_URL}/", wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(3000)

        inputs = page.locator("input:visible").all()
        if len(inputs) >= 1:
            inputs[0].fill(usuario)
        if len(inputs) >= 2:
            inputs[1].fill(password)
        else:
            try:
                page.locator("input[type='password']").first.fill(password)
            except Exception:
                pass

        try:
            page.locator("button[type='submit']").click(timeout=8000)
        except Exception:
            try:
                page.locator(
                    "button:has-text('Iniciar'), button:has-text('Login'), button:has-text('Ingresar')"
                ).first.click(timeout=5000)
            except Exception:
                try:
                    page.locator("button").first.click(timeout=8000)
                except Exception:
                    page.keyboard.press("Enter")

        page.wait_for_timeout(6000)

        if "/login" not in page.url:
            return

        if intento < max_intentos:
            page.wait_for_timeout(5000)
        else:
            raise Exception(f"Login C21 fallido tras {max_intentos} intentos — URL: {page.url}")


def exportar_reporte(page, fecha_desde: str, fecha_hasta: str) -> Path:
    """
    Usa el botón nativo "Descargar" (verde, debajo de los filtros) —
    confirmado que exporta el Excel completo con todas las columnas
    (PRECIO CIERRE, M2T, M2C, LAT, LON, etc.), sin necesitar ningún script
    de consola.

    Selector confirmado contra el sitio real: es un <a>, no un <button>
    (por eso get_by_role("button", ...) no lo encontraba).
    """
    url = construir_url_busqueda(fecha_desde, fecha_hasta)
    # La página puede tardar bastante (confirmado: hasta ~110 segundos,
    # con el ícono de carga girando en la pestaña) y a veces hace una
    # recarga automática antes de mostrar los resultados finales.
    # Timeout generoso a propósito (150s), con margen extra sobre lo medido.
    page.goto(url, wait_until="networkidle", timeout=150000)
    page.wait_for_selector("#myTableBusqueda tbody tr", timeout=150000)

    with page.expect_download(timeout=60000) as descarga_info:
        page.locator("a.btnDescargarInventarioExcel").click()
    descarga = descarga_info.value
    descarga.save_as(DESCARGA_TMP)
    return DESCARGA_TMP


def parse_money_usd(s):
    if not s:
        return None
    s = str(s)
    es_pyg = "PYG" in s.upper()
    s = s.replace("USD", "").replace("PYG", "").replace("Gs.", "").replace("Gs", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        val = float(s)
    except ValueError:
        return None
    if es_pyg:
        val = val / TC_APROX
    return val


def to_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def parsear_excel(path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active  # el nombre de la hoja varia segun la exportacion (ej. "Hoja1")
    rows = list(ws.iter_rows(values_only=True))
    header = rows[0]
    idx = {h: i for i, h in enumerate(header) if h}

    out = []
    for r in rows[1:]:
        precio = parse_money_usd(r[idx["PRECIO CIERRE"]])
        precio_lista = parse_money_usd(r[idx["PRECIO"]])
        m2t = to_float(r[idx["M2T"]])
        m2c = to_float(r[idx["M2C"]])
        fc = r[idx["FECHA CIERRE"]]
        fa = r[idx["FECHA ALTA"]]
        pid = r[idx["ID"]]
        lat = to_float(r[idx["LAT"]])
        lon = to_float(r[idx["LON"]])
        if precio is None or not m2t or m2t <= 0 or not fc or not pid:
            continue
        out.append({
            "id": int(pid),
            "mu": r[idx["MUNICIPIO"]],
            "ti": r[idx["TIPO"]],
            "su": r[idx["SUBTIPO"]],
            "dir": r[idx["DIRECCION"]],
            "pr": round(precio, 2),
            "prL": round(precio_lista, 2) if precio_lista else None,
            "mt": m2t,
            "mc": m2c,
            "fc": str(fc)[:10],
            "fa": str(fa)[:10] if fa else None,
            "la": lat,
            "lo": lon,
        })
    return out


def cargar_dataset_actual() -> list[dict]:
    if DATASET_PATH.exists():
        with open(DATASET_PATH, encoding="utf-8") as f:
            return json.load(f)
    return []


def fecha_de_corte(dataset: list[dict]) -> str:
    if not dataset:
        # sin dataset previo: primer corrida, usar un rango amplio por defecto
        return (datetime.now() - timedelta(days=548)).strftime("%Y-%m-%d")  # ~18 meses
    ultima = max(d["fc"] for d in dataset)
    corte = datetime.strptime(ultima, "%Y-%m-%d") - timedelta(days=MARGEN_DIAS)
    return corte.strftime("%Y-%m-%d")


def upsert(dataset: list[dict], nuevos: list[dict]) -> tuple[list[dict], int, int]:
    por_id = {d["id"]: d for d in dataset}
    agregados = 0
    actualizados = 0
    for n in nuevos:
        if n["id"] in por_id:
            if por_id[n["id"]] != n:
                actualizados += 1
            por_id[n["id"]] = n
        else:
            por_id[n["id"]] = n
            agregados += 1
    resultado = list(por_id.values())
    resultado.sort(key=lambda d: d["fc"])
    return resultado, agregados, actualizados


def git_commit_push(mensaje: str):
    subprocess.run(["git", "add", "data/cerradas.json"], cwd=REPO_DIR, check=True)
    resultado = subprocess.run(["git", "diff", "--cached", "--quiet"], cwd=REPO_DIR)
    if resultado.returncode == 0:
        print("Sin cambios para commitear.")
        return
    subprocess.run(["git", "commit", "-m", mensaje], cwd=REPO_DIR, check=True)
    subprocess.run(["git", "push"], cwd=REPO_DIR, check=True)


def main():
    dataset = cargar_dataset_actual()
    fecha_desde = fecha_de_corte(dataset)
    fecha_hasta = datetime.now().strftime("%Y-%m-%d")
    print(f"Dataset actual: {len(dataset)} registros. Exportando desde {fecha_desde} hasta {fecha_hasta}.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        login(page)  # <-- reemplazar por login real
        ruta_export = exportar_reporte(page, fecha_desde, fecha_hasta)  # <-- completar selectores reales
        browser.close()

    nuevos = parsear_excel(ruta_export)
    print(f"Exportados {len(nuevos)} registros en el rango.")

    dataset_final, agregados, actualizados = upsert(dataset, nuevos)
    DATASET_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(DATASET_PATH, "w", encoding="utf-8") as f:
        json.dump(dataset_final, f, ensure_ascii=False, separators=(",", ":"))

    print(f"Agregados: {agregados} | Actualizados: {actualizados} | Total dataset: {len(dataset_final)}")

    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    git_commit_push(f"Actualización semanal cerradas {fecha_hoy}: +{agregados} nuevos, {actualizados} actualizados")

    ruta_export.unlink(missing_ok=True)


if __name__ == "__main__":
    main()