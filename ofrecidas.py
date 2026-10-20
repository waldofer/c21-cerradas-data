"""Propiedades ofrecidas (a la venta) exportadas de 21online.

Modulo puro: solo openpyxl y la libreria estandar. No importa
`actualizar_cerradas` (ese arrastra playwright).

Privacidad: el export trae datos del propietario y columnas comerciales. Nada de
eso puede salir de aca; la salida se arma SIEMPRE desde la lista blanca `CAMPOS`.
"""
import re
from pathlib import Path

import openpyxl

# Orden y lista blanca de claves de salida.
CAMPOS = ("id", "mu", "ti", "su", "dir", "en", "of", "pp", "mon", "mt", "mc", "fa", "la", "lo")

COLUMNAS_REQUERIDAS = {"ID", "PRECIO", "M2T", "FECHA ALTA", "MUNICIPIO"}

# Columna del export -> clave de salida (texto sin conversion).
_TEXTOS = {
    "mu": "MUNICIPIO",
    "ti": "TIPO",
    "su": "SUBTIPO",
    "dir": "DIRECCION",
    "en": "ENCABEZADO",
    "of": "OFICINA",
}

# Moneda: "USD", "PYG" o "Gs"/"Gs." (guaranies), antes o despues del numero.
_RE_MONEDA = re.compile(r"\b(USD|PYG|GS)\b\.?", re.IGNORECASE)


def parse_precio(s):
    """'420.000.000,00 PYG' -> (420000000.0, 'PYG'). None si no hay numero o moneda."""
    if s is None:
        return None
    s = str(s).strip()
    m = _RE_MONEDA.search(s)
    if not m:
        return None
    moneda = "USD" if m.group(1).upper() == "USD" else "PYG"
    numero = (s[:m.start()] + " " + s[m.end():]).strip()
    # Formato paraguayo: punto de miles, coma decimal.
    numero = numero.replace(".", "").replace(",", ".")
    try:
        return float(numero), moneda
    except ValueError:
        return None


def validar_columnas(header):
    """Motivo (str) si falta alguna columna requerida; None si esta todo."""
    presentes = {str(h).strip() for h in header if h}
    faltantes = COLUMNAS_REQUERIDAS - presentes
    if faltantes:
        return f"Faltan columnas esperadas en el Excel de ofrecidas: {sorted(faltantes)}"
    return None


def _texto(v):
    """Celda -> str sin espacios, o None si esta vacia."""
    if v is None:
        return None
    v = str(v).strip()
    return v or None


def _numero(v):
    """Celda de texto/numero -> float, o None si esta vacia o no es un numero."""
    try:
        return float(_texto(v))
    except (TypeError, ValueError):
        return None


def parsear_ofrecidas(path: Path) -> list[dict]:
    """Un dict por ID (gana la ultima fila), solo con las claves de CAMPOS.

    Se descartan las filas sin ID, con precio ilegible o con M2T faltante o <= 0.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active  # el nombre de la hoja varia segun la exportacion
        filas = ws.iter_rows(values_only=True)
        header = next(filas, None) or ()
        motivo = validar_columnas(header)
        if motivo:
            raise ValueError(motivo)
        idx = {str(h).strip(): i for i, h in enumerate(header) if h}

        def celda(fila, col):
            i = idx.get(col)
            return fila[i] if i is not None and i < len(fila) else None

        por_id = {}
        for fila in filas:
            pid = _numero(celda(fila, "ID"))
            precio = parse_precio(celda(fila, "PRECIO"))
            m2t = _numero(celda(fila, "M2T"))
            if pid is None or precio is None or m2t is None or m2t <= 0:
                continue
            fa = _texto(celda(fila, "FECHA ALTA"))
            valores = {
                "id": int(pid),
                "pp": precio[0],
                "mon": precio[1],
                "mt": m2t,
                "mc": _numero(celda(fila, "M2C")),
                "fa": fa[:10] if fa else None,
                "la": _numero(celda(fila, "LAT")),
                "lo": _numero(celda(fila, "LON")),
            }
            for clave, col in _TEXTOS.items():
                valores[clave] = _texto(celda(fila, col))
            # Lista blanca: solo salen las claves de CAMPOS.
            por_id[valores["id"]] = {k: valores[k] for k in CAMPOS}
        return list(por_id.values())
    finally:
        wb.close()


# Si la descarga trae menos del 80% de las ofrecidas de la semana anterior,
# se la trata como rota y no se publica.
UMBRAL_DESCARGA = 0.8


def descarga_sospechosa(n_nuevas: int, n_anteriores: int, umbral: float = UMBRAL_DESCARGA):
    """Mensaje (str) si la descarga parece cortada; None si es normal o es la primera corrida."""
    if n_anteriores == 0 or n_nuevas >= umbral * n_anteriores:
        return None
    return (
        f"La descarga de ofrecidas trajo {n_nuevas} propiedades y la semana anterior "
        f"habia {n_anteriores} (menos del {umbral:.0%}); se la considera incompleta."
    )


def comparar(anteriores: list[dict], descargadas: list[dict], vendidas_por_id: dict[int, dict]) -> dict:
    """Compara la descarga de esta semana con la anterior y con las vendidas.

    Una ofrecida que ya figura en `vendidas_por_id` no se publica como ofrecida.
    """
    ant = {r["id"]: r for r in anteriores}
    desc = {r["id"]: r for r in descargadas}

    ofrecidas = sorted(
        (r for i, r in desc.items() if i not in vendidas_por_id), key=lambda r: r["id"]
    )
    nuevas = [r for r in ofrecidas if r["id"] not in ant]
    pasaron_a_vendidas = [
        {
            "id": i,
            "dir": r["dir"],
            "mu": r["mu"],
            "pp": r["pp"],
            "mon": r["mon"],
            "pr_cierre": vendidas_por_id[i]["pr"],
        }
        for i, r in ant.items()
        if i in vendidas_por_id
    ]
    dadas_de_baja = [
        r for i, r in ant.items() if i not in desc and i not in vendidas_por_id
    ]
    cambios_precio = [
        {
            "id": i,
            "antes": (ant[i]["pp"], ant[i]["mon"]),
            "despues": (r["pp"], r["mon"]),
        }
        for i, r in desc.items()
        if i in ant and (ant[i]["pp"], ant[i]["mon"]) != (r["pp"], r["mon"])
    ]
    return {
        "ofrecidas": ofrecidas,
        "nuevas": nuevas,
        "pasaron_a_vendidas": pasaron_a_vendidas,
        "dadas_de_baja": dadas_de_baja,
        "cambios_precio": cambios_precio,
        "en_venta_antes": len(anteriores),
    }
