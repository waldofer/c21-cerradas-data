"""
Serie histórica de cotizaciones Gs/USD del Banco Central del Paraguay.

El BCP no publica API REST. El único acceso programático es este endpoint de
export, que a pesar del sufijo .xls devuelve una TABLA HTML servida con
Content-Type de Excel. Se parsea con regex, sin librerías de Excel.

Uso:
    python cotizaciones_bcp.py            # genera data/cotizaciones_bcp.json
    python cotizaciones_bcp.py --refrescar-anho 2026
"""

import argparse
import calendar
import json
import re
import sys
from pathlib import Path

import requests

REPO_DIR = Path(__file__).parent
TABLA_PATH = REPO_DIR / "data" / "cotizaciones_bcp.json"

URL_BCP = "https://www.bcp.gov.py/webapps/web/cotizacion/monedas-historica/xls"
# Sin User-Agent de navegador el BCP responde 403.
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    )
}
MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
         "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]

ANHO_DESDE = 2019
ANHO_HASTA = 2026

# Fechas de control: si alguna no coincide, el parseo está mal.
VERIFICACION = {
    "2019-01-15": 6043.98,
    "2020-04-20": 6483.08,
    "2021-03-15": 6618.42,
    "2022-06-10": 6854.08,
    "2023-01-20": 7439.13,
}


def _num_py(txt):
    """'6.016,59' -> 6016.59 ; 'ND' o vacío -> None"""
    txt = txt.strip()
    if not txt or txt.upper() == "ND":
        return None
    try:
        return float(txt.replace(".", "").replace(",", "."))
    except ValueError:
        return None


def descargar_anho(anho: int) -> dict:
    """Descarga y parsea un año. Devuelve {'YYYY-MM-DD': float}."""
    resp = requests.post(
        URL_BCP, data={"anho": str(anho), "moneda": "USD"},
        headers=HEADERS, timeout=60,
    )
    resp.raise_for_status()
    html = resp.text

    filas = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.S | re.I)
    if not filas:
        raise ValueError(f"{anho}: la respuesta no contiene filas de tabla")

    # Mapear columna -> mes leyendo el encabezado, en vez de asumir el orden.
    col_a_mes = {}
    for fila in filas:
        celdas = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", fila, re.S | re.I)
        textos = [re.sub(r"<[^>]+>", "", c).strip().upper() for c in celdas]
        encontrados = {i: MESES.index(t) + 1 for i, t in enumerate(textos) if t in MESES}
        if len(encontrados) == 12:
            col_a_mes = encontrados
            break
    if not col_a_mes:
        raise ValueError(f"{anho}: no se encontró la fila de encabezado con los 12 meses")

    out = {}
    for fila in filas:
        celdas = re.findall(r"<t[hd][^>]*>(.*?)</t[hd]>", fila, re.S | re.I)
        if not celdas:
            continue
        textos = [re.sub(r"<[^>]+>", "", c).strip() for c in celdas]
        if not textos[0].isdigit():
            continue
        dia = int(textos[0])
        for col, mes in col_a_mes.items():
            if col >= len(textos):
                continue
            valor = _num_py(textos[col])
            if valor is None:
                continue
            # Descartar fechas inexistentes (31 de febrero, etc.)
            if dia > calendar.monthrange(anho, mes)[1]:
                continue
            out[f"{anho}-{mes:02d}-{dia:02d}"] = valor
    return out


def construir_tabla(desde: int = ANHO_DESDE, hasta: int = ANHO_HASTA) -> dict:
    tabla = {}
    for anho in range(desde, hasta + 1):
        parcial = descargar_anho(anho)
        print(f"  {anho}: {len(parcial)} días")
        tabla.update(parcial)
    return dict(sorted(tabla.items()))


def cotizacion_de(tabla: dict, fecha: str, max_retroceso: int = 10):
    """
    Cotización de una fecha, retrocediendo al día hábil anterior si hace falta.
    Devuelve (cotizacion, dias_retrocedidos) o (None, -1) si no se resolvió.
    """
    from datetime import datetime, timedelta
    d = datetime.strptime(fecha, "%Y-%m-%d")
    for i in range(max_retroceso + 1):
        clave = (d - timedelta(days=i)).strftime("%Y-%m-%d")
        if clave in tabla:
            return tabla[clave], i
    return None, -1


def verificar(tabla: dict) -> list:
    """Devuelve la lista de discrepancias contra las fechas de control."""
    fallos = []
    for fecha, esperado in VERIFICACION.items():
        real = tabla.get(fecha)
        if real != esperado:
            fallos.append(f"{fecha}: esperado {esperado}, obtenido {real}")
    return fallos


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--refrescar-anho", type=int, default=None,
                    help="Refresca solo ese año sobre la tabla ya existente")
    args = ap.parse_args()

    if args.refrescar_anho:
        tabla = json.loads(TABLA_PATH.read_text(encoding="utf-8")) if TABLA_PATH.exists() else {}
        print(f"Refrescando {args.refrescar_anho}…")
        tabla.update(descargar_anho(args.refrescar_anho))
        tabla = dict(sorted(tabla.items()))
    else:
        print(f"Descargando BCP {ANHO_DESDE}-{ANHO_HASTA}…")
        tabla = construir_tabla()

    fallos = verificar(tabla)
    if fallos:
        print("VERIFICACIÓN FALLIDA — no se escribe el archivo:", file=sys.stderr)
        for f in fallos:
            print("  " + f, file=sys.stderr)
        sys.exit(1)

    TABLA_PATH.parent.mkdir(parents=True, exist_ok=True)
    TABLA_PATH.write_text(json.dumps(tabla, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"OK: {len(tabla)} días en {TABLA_PATH}")
    print(f"Rango: {min(tabla)} a {max(tabla)}")


if __name__ == "__main__":
    main()
