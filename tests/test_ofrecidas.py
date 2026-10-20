import json

import openpyxl

from ofrecidas import (
    CAMPOS,
    COLUMNAS_REQUERIDAS,
    parse_precio,
    parsear_ofrecidas,
    validar_columnas,
)

# Encabezado real del export de 21online (el orden es el del sistema).
HEADER = [
    "LICENCIA", "OFICINA", "FHE", "COMERCIAL", "ID", "CLAVE", "ESTATUS", "DIRECCION",
    "COLONIA", "MUNICIPIO", "ESTADO", "PAIS", "DESARROLLO", "FECHA ALTA",
    "FECHA PROMOCIÓN", "TIPO", "SUBTIPO", "OPERACIÓN", "PRECIO", "TIPO CONTRATO",
    "DIAS CONTRATO", "PRECIO CONTRATO", "TIPO DE COMISION", "COMISION CONTRATO",
    "COMISION A COMPARTIR", "ANTIGUEDAD", "PISO", "HABITACIONES", "M2T", "M2C",
    "LAT", "LON", "INTERNET", "ASESOR EXCLUSIVA", "PROPIETARIO", "TELEFONO PROPIETARIO",
    "EMAIL PROPIETARIO", "FECHA SUSPENDIDA", "FECHA INTENCIÓN", "FECHA CANCELACIÓN",
    "FECHA CIERRE", "PRECIO CIERRE", "COMISIÓN CIERRE", "COMISIÓN CIERRE VENDEDOR",
    "COMISIÓN CIERRE COMPRADOR", "ASESOR CIERRE", "COMPARTIDA", "COMPARTIDA CON",
    "Contraparte", "COMISION COMPARTIDA", "REFERIDO", "REFERIDO CON", "CANTIDAD REFERIDA",
    "MES CIERRE", "ZONA", "ENCABEZADO", "NOTAS",
]

PROPIETARIO = "Juan Perez"
TELEFONO = "0981123456"
EMAIL = "juan@example.com"


def fila_valida(**cambios):
    """Una fila valida con todas las celdas como texto, igual que el export real."""
    fila = {
        "OFICINA": "CENTURY 21 Demo",
        "ID": "51505",
        "DIRECCION": "Calle Falsa 123",
        "MUNICIPIO": "Asuncion",
        "FECHA ALTA": "2026-07-17",
        "TIPO": "Casa",
        "SUBTIPO": "Casa",
        "PRECIO": "420.000.000,00 PYG",
        "M2T": "360",
        "M2C": "200",
        "LAT": "-25.26",
        "LON": "-57.58",
        "PROPIETARIO": PROPIETARIO,
        "TELEFONO PROPIETARIO": TELEFONO,
        "EMAIL PROPIETARIO": EMAIL,
        "COMISION CONTRATO": "3",
        "NOTAS": "Llamar al propietario",
        "ENCABEZADO": "Casa en Asuncion",
    }
    fila.update(cambios)
    return fila


def armar_excel(tmp_path, filas, columnas_extra=()):
    """Escribe un xlsx con el encabezado real (+ columnas extra); celdas ausentes quedan vacias."""
    header = HEADER + list(columnas_extra)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for fila in filas:
        ws.append([fila.get(h, "") for h in header])
    path = tmp_path / "ofrecidas.xlsx"
    wb.save(path)
    return path


def test_parse_precio_pyg():
    assert parse_precio("420.000.000,00 PYG") == (420000000.0, "PYG")


def test_parse_precio_usd_sufijo_y_prefijo():
    assert parse_precio("150.000,00 USD") == (150000.0, "USD")
    assert parse_precio("USD 150.000") == (150000.0, "USD")


def test_parse_precio_chico_se_acepta():
    assert parse_precio("50,00 USD") == (50.0, "USD")


def test_parse_precio_ilegible():
    for x in (None, "", "Consultar", "150.000"):
        assert parse_precio(x) is None


def test_parse_precio_gs_es_pyg():
    assert parse_precio("Gs. 1.500.000") == (1500000.0, "PYG")
    assert parse_precio("350.000.000 Gs") == (350000000.0, "PYG")


def test_parsear_solo_campos_permitidos(tmp_path):
    filas = [
        fila_valida(**{"CELULAR PROPIETARIO": "0971999888"}),
        fila_valida(ID="51506"),
    ]
    path = armar_excel(tmp_path, filas, columnas_extra=("CELULAR PROPIETARIO",))
    resultado = parsear_ofrecidas(path)
    assert len(resultado) == 2
    for r in resultado:
        assert set(r) == set(CAMPOS)
    volcado = json.dumps(resultado)
    for dato in (PROPIETARIO, TELEFONO, EMAIL, "0971999888", "Llamar al propietario"):
        assert dato not in volcado


def test_parsear_descarta_incompletas(tmp_path):
    filas = [
        fila_valida(ID=""),
        fila_valida(ID="1", PRECIO="Consultar"),
        fila_valida(ID="2", M2T=""),
        fila_valida(ID="3", M2T="0"),
    ]
    assert parsear_ofrecidas(armar_excel(tmp_path, filas)) == []

    r = parsear_ofrecidas(armar_excel(tmp_path, [fila_valida()]))
    assert len(r) == 1
    assert r[0]["id"] == 51505
    assert r[0]["pp"] == 420000000.0
    assert r[0]["mon"] == "PYG"
    assert r[0]["mt"] == 360.0
    assert r[0]["mc"] == 200.0
    assert r[0]["fa"] == "2026-07-17"
    assert r[0]["la"] == -25.26
    assert r[0]["lo"] == -57.58
    assert r[0]["mu"] == "Asuncion"
    assert r[0]["of"] == "CENTURY 21 Demo"
    assert r[0]["en"] == "Casa en Asuncion"


def test_parsear_celdas_vacias_son_none(tmp_path):
    fila = fila_valida(M2C="", LAT="", LON="", **{"FECHA ALTA": "", "ENCABEZADO": ""})
    (r,) = parsear_ofrecidas(armar_excel(tmp_path, [fila]))
    assert r["mc"] is None and r["la"] is None and r["lo"] is None
    assert r["fa"] is None and r["en"] is None


def test_parsear_id_duplicado(tmp_path):
    filas = [
        fila_valida(ID="51505", PRECIO="100.000,00 USD"),
        fila_valida(ID="51505", PRECIO="120.000,00 USD"),
    ]
    r = parsear_ofrecidas(armar_excel(tmp_path, filas))
    assert len(r) == 1
    assert r[0]["pp"] == 120000.0


def test_validar_columnas():
    sin_precio = [h for h in HEADER if h != "PRECIO"]
    assert "PRECIO" in validar_columnas(sin_precio)
    assert validar_columnas(HEADER) is None
    assert COLUMNAS_REQUERIDAS <= set(HEADER)
